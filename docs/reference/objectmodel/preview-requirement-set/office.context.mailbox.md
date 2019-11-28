---
title: Office のメールボックス-プレビュー要件セット
description: ''
ms.date: 11/25/2019
localization_priority: Normal
ms.openlocfilehash: 8c67f7cf9231dd1c0db0d9a8d4ae9fb48e458435
ms.sourcegitcommit: 05a883a7fd89136301ce35aabc57638e9f563288
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/27/2019
ms.locfileid: "39629196"
---
# <a name="mailbox"></a><span data-ttu-id="e3f6b-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="e3f6b-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="e3f6b-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="e3f6b-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="e3f6b-104">Microsoft Outlook の Outlook アドイン オブジェクト モデルへのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="e3f6b-105">要件</span><span class="sxs-lookup"><span data-stu-id="e3f6b-105">Requirements</span></span>

|<span data-ttu-id="e3f6b-106">要件</span><span class="sxs-lookup"><span data-stu-id="e3f6b-106">Requirement</span></span>| <span data-ttu-id="e3f6b-107">値</span><span class="sxs-lookup"><span data-stu-id="e3f6b-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="e3f6b-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e3f6b-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e3f6b-109">1.0</span><span class="sxs-lookup"><span data-stu-id="e3f6b-109">1.0</span></span>|
|[<span data-ttu-id="e3f6b-110">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e3f6b-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e3f6b-111">制限あり</span><span class="sxs-lookup"><span data-stu-id="e3f6b-111">Restricted</span></span>|
|[<span data-ttu-id="e3f6b-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e3f6b-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e3f6b-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e3f6b-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="e3f6b-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="e3f6b-114">Properties</span></span>

| <span data-ttu-id="e3f6b-115">プロパティ</span><span class="sxs-lookup"><span data-stu-id="e3f6b-115">Property</span></span> | <span data-ttu-id="e3f6b-116">最小値</span><span class="sxs-lookup"><span data-stu-id="e3f6b-116">Minimum</span></span><br><span data-ttu-id="e3f6b-117">アクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e3f6b-117">permission level</span></span> | <span data-ttu-id="e3f6b-118">モード</span><span class="sxs-lookup"><span data-stu-id="e3f6b-118">Modes</span></span> | <span data-ttu-id="e3f6b-119">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="e3f6b-119">Return type</span></span> | <span data-ttu-id="e3f6b-120">最小値</span><span class="sxs-lookup"><span data-stu-id="e3f6b-120">Minimum</span></span><br><span data-ttu-id="e3f6b-121">要件セット</span><span class="sxs-lookup"><span data-stu-id="e3f6b-121">requirement set</span></span> |
|---|---|---|---|---|
| [<span data-ttu-id="e3f6b-122">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="e3f6b-122">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="e3f6b-123">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e3f6b-123">ReadItem</span></span> | <span data-ttu-id="e3f6b-124">作成</span><span class="sxs-lookup"><span data-stu-id="e3f6b-124">Compose</span></span><br><span data-ttu-id="e3f6b-125">読み取り</span><span class="sxs-lookup"><span data-stu-id="e3f6b-125">Read</span></span> | <span data-ttu-id="e3f6b-126">String</span><span class="sxs-lookup"><span data-stu-id="e3f6b-126">String</span></span> | <span data-ttu-id="e3f6b-127">1.0</span><span class="sxs-lookup"><span data-stu-id="e3f6b-127">1.0</span></span> |
| [<span data-ttu-id="e3f6b-128">masterCategories</span><span class="sxs-lookup"><span data-stu-id="e3f6b-128">masterCategories</span></span>](#mastercategories-mastercategories) | <span data-ttu-id="e3f6b-129">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="e3f6b-129">ReadWriteMailbox</span></span> | <span data-ttu-id="e3f6b-130">作成</span><span class="sxs-lookup"><span data-stu-id="e3f6b-130">Compose</span></span><br><span data-ttu-id="e3f6b-131">読み取り</span><span class="sxs-lookup"><span data-stu-id="e3f6b-131">Read</span></span> | [<span data-ttu-id="e3f6b-132">MasterCategories</span><span class="sxs-lookup"><span data-stu-id="e3f6b-132">MasterCategories</span></span>](/javascript/api/outlook/office.mastercategories) | <span data-ttu-id="e3f6b-133">プレビュー</span><span class="sxs-lookup"><span data-stu-id="e3f6b-133">Preview</span></span> |
| [<span data-ttu-id="e3f6b-134">restUrl</span><span class="sxs-lookup"><span data-stu-id="e3f6b-134">restUrl</span></span>](#resturl-string) | <span data-ttu-id="e3f6b-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e3f6b-135">ReadItem</span></span> | <span data-ttu-id="e3f6b-136">作成</span><span class="sxs-lookup"><span data-stu-id="e3f6b-136">Compose</span></span><br><span data-ttu-id="e3f6b-137">読み取り</span><span class="sxs-lookup"><span data-stu-id="e3f6b-137">Read</span></span> | <span data-ttu-id="e3f6b-138">String</span><span class="sxs-lookup"><span data-stu-id="e3f6b-138">String</span></span> | <span data-ttu-id="e3f6b-139">1.5</span><span class="sxs-lookup"><span data-stu-id="e3f6b-139">1.5</span></span> |

##### <a name="methods"></a><span data-ttu-id="e3f6b-140">メソッド</span><span class="sxs-lookup"><span data-stu-id="e3f6b-140">Methods</span></span>

| <span data-ttu-id="e3f6b-141">メソッド</span><span class="sxs-lookup"><span data-stu-id="e3f6b-141">Method</span></span> | <span data-ttu-id="e3f6b-142">最小値</span><span class="sxs-lookup"><span data-stu-id="e3f6b-142">Minimum</span></span><br><span data-ttu-id="e3f6b-143">アクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e3f6b-143">permission level</span></span> | <span data-ttu-id="e3f6b-144">モード</span><span class="sxs-lookup"><span data-stu-id="e3f6b-144">Modes</span></span> | <span data-ttu-id="e3f6b-145">最小値</span><span class="sxs-lookup"><span data-stu-id="e3f6b-145">Minimum</span></span><br><span data-ttu-id="e3f6b-146">要件セット</span><span class="sxs-lookup"><span data-stu-id="e3f6b-146">requirement set</span></span> |
|---|---|---|---|
| [<span data-ttu-id="e3f6b-147">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="e3f6b-147">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="e3f6b-148">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e3f6b-148">ReadItem</span></span> | <span data-ttu-id="e3f6b-149">作成</span><span class="sxs-lookup"><span data-stu-id="e3f6b-149">Compose</span></span><br><span data-ttu-id="e3f6b-150">読み取り</span><span class="sxs-lookup"><span data-stu-id="e3f6b-150">Read</span></span> | <span data-ttu-id="e3f6b-151">1.5</span><span class="sxs-lookup"><span data-stu-id="e3f6b-151">1.5</span></span> |
| [<span data-ttu-id="e3f6b-152">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="e3f6b-152">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="e3f6b-153">制限あり</span><span class="sxs-lookup"><span data-stu-id="e3f6b-153">Restricted</span></span> | <span data-ttu-id="e3f6b-154">作成</span><span class="sxs-lookup"><span data-stu-id="e3f6b-154">Compose</span></span><br><span data-ttu-id="e3f6b-155">読み取り</span><span class="sxs-lookup"><span data-stu-id="e3f6b-155">Read</span></span> | <span data-ttu-id="e3f6b-156">1.3</span><span class="sxs-lookup"><span data-stu-id="e3f6b-156">1.3</span></span> |
| [<span data-ttu-id="e3f6b-157">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="e3f6b-157">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="e3f6b-158">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e3f6b-158">ReadItem</span></span> | <span data-ttu-id="e3f6b-159">作成</span><span class="sxs-lookup"><span data-stu-id="e3f6b-159">Compose</span></span><br><span data-ttu-id="e3f6b-160">読み取り</span><span class="sxs-lookup"><span data-stu-id="e3f6b-160">Read</span></span> | <span data-ttu-id="e3f6b-161">1.0</span><span class="sxs-lookup"><span data-stu-id="e3f6b-161">1.0</span></span> |
| [<span data-ttu-id="e3f6b-162">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="e3f6b-162">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="e3f6b-163">制限あり</span><span class="sxs-lookup"><span data-stu-id="e3f6b-163">Restricted</span></span> | <span data-ttu-id="e3f6b-164">作成</span><span class="sxs-lookup"><span data-stu-id="e3f6b-164">Compose</span></span><br><span data-ttu-id="e3f6b-165">読み取り</span><span class="sxs-lookup"><span data-stu-id="e3f6b-165">Read</span></span> | <span data-ttu-id="e3f6b-166">1.3</span><span class="sxs-lookup"><span data-stu-id="e3f6b-166">1.3</span></span> |
| [<span data-ttu-id="e3f6b-167">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="e3f6b-167">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="e3f6b-168">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e3f6b-168">ReadItem</span></span> | <span data-ttu-id="e3f6b-169">作成</span><span class="sxs-lookup"><span data-stu-id="e3f6b-169">Compose</span></span><br><span data-ttu-id="e3f6b-170">読み取り</span><span class="sxs-lookup"><span data-stu-id="e3f6b-170">Read</span></span> | <span data-ttu-id="e3f6b-171">1.0</span><span class="sxs-lookup"><span data-stu-id="e3f6b-171">1.0</span></span> |
| [<span data-ttu-id="e3f6b-172">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="e3f6b-172">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="e3f6b-173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e3f6b-173">ReadItem</span></span> | <span data-ttu-id="e3f6b-174">作成</span><span class="sxs-lookup"><span data-stu-id="e3f6b-174">Compose</span></span><br><span data-ttu-id="e3f6b-175">読み取り</span><span class="sxs-lookup"><span data-stu-id="e3f6b-175">Read</span></span> | <span data-ttu-id="e3f6b-176">1.0</span><span class="sxs-lookup"><span data-stu-id="e3f6b-176">1.0</span></span> |
| [<span data-ttu-id="e3f6b-177">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="e3f6b-177">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="e3f6b-178">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e3f6b-178">ReadItem</span></span> | <span data-ttu-id="e3f6b-179">作成</span><span class="sxs-lookup"><span data-stu-id="e3f6b-179">Compose</span></span><br><span data-ttu-id="e3f6b-180">読み取り</span><span class="sxs-lookup"><span data-stu-id="e3f6b-180">Read</span></span> | <span data-ttu-id="e3f6b-181">1.0</span><span class="sxs-lookup"><span data-stu-id="e3f6b-181">1.0</span></span> |
| [<span data-ttu-id="e3f6b-182">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="e3f6b-182">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="e3f6b-183">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e3f6b-183">ReadItem</span></span> | <span data-ttu-id="e3f6b-184">読み取り</span><span class="sxs-lookup"><span data-stu-id="e3f6b-184">Read</span></span> | <span data-ttu-id="e3f6b-185">1.0</span><span class="sxs-lookup"><span data-stu-id="e3f6b-185">1.0</span></span> |
| [<span data-ttu-id="e3f6b-186">displayNewMessageForm</span><span class="sxs-lookup"><span data-stu-id="e3f6b-186">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="e3f6b-187">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e3f6b-187">ReadItem</span></span> | <span data-ttu-id="e3f6b-188">作成</span><span class="sxs-lookup"><span data-stu-id="e3f6b-188">Compose</span></span><br><span data-ttu-id="e3f6b-189">読み取り</span><span class="sxs-lookup"><span data-stu-id="e3f6b-189">Read</span></span> | <span data-ttu-id="e3f6b-190">1.6</span><span class="sxs-lookup"><span data-stu-id="e3f6b-190">1.6</span></span> |
| [<span data-ttu-id="e3f6b-191">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="e3f6b-191">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="e3f6b-192">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e3f6b-192">ReadItem</span></span> | <span data-ttu-id="e3f6b-193">作成</span><span class="sxs-lookup"><span data-stu-id="e3f6b-193">Compose</span></span><br><span data-ttu-id="e3f6b-194">読み取り</span><span class="sxs-lookup"><span data-stu-id="e3f6b-194">Read</span></span> | <span data-ttu-id="e3f6b-195">1.5</span><span class="sxs-lookup"><span data-stu-id="e3f6b-195">1.5</span></span> |
| [<span data-ttu-id="e3f6b-196">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="e3f6b-196">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="e3f6b-197">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e3f6b-197">ReadItem</span></span> | <span data-ttu-id="e3f6b-198">作成</span><span class="sxs-lookup"><span data-stu-id="e3f6b-198">Compose</span></span><br><span data-ttu-id="e3f6b-199">読み取り</span><span class="sxs-lookup"><span data-stu-id="e3f6b-199">Read</span></span> | <span data-ttu-id="e3f6b-200">1.3</span><span class="sxs-lookup"><span data-stu-id="e3f6b-200">1.3</span></span><br><span data-ttu-id="e3f6b-201">1.0</span><span class="sxs-lookup"><span data-stu-id="e3f6b-201">1.0</span></span> |
| [<span data-ttu-id="e3f6b-202">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="e3f6b-202">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="e3f6b-203">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e3f6b-203">ReadItem</span></span> | <span data-ttu-id="e3f6b-204">作成</span><span class="sxs-lookup"><span data-stu-id="e3f6b-204">Compose</span></span><br><span data-ttu-id="e3f6b-205">読み取り</span><span class="sxs-lookup"><span data-stu-id="e3f6b-205">Read</span></span> | <span data-ttu-id="e3f6b-206">1.0</span><span class="sxs-lookup"><span data-stu-id="e3f6b-206">1.0</span></span> |
| [<span data-ttu-id="e3f6b-207">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="e3f6b-207">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="e3f6b-208">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="e3f6b-208">ReadWriteMailbox</span></span> | <span data-ttu-id="e3f6b-209">作成</span><span class="sxs-lookup"><span data-stu-id="e3f6b-209">Compose</span></span><br><span data-ttu-id="e3f6b-210">読み取り</span><span class="sxs-lookup"><span data-stu-id="e3f6b-210">Read</span></span> | <span data-ttu-id="e3f6b-211">1.0</span><span class="sxs-lookup"><span data-stu-id="e3f6b-211">1.0</span></span> |
| [<span data-ttu-id="e3f6b-212">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="e3f6b-212">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="e3f6b-213">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e3f6b-213">ReadItem</span></span> | <span data-ttu-id="e3f6b-214">作成</span><span class="sxs-lookup"><span data-stu-id="e3f6b-214">Compose</span></span><br><span data-ttu-id="e3f6b-215">読み取り</span><span class="sxs-lookup"><span data-stu-id="e3f6b-215">Read</span></span> | <span data-ttu-id="e3f6b-216">1.5</span><span class="sxs-lookup"><span data-stu-id="e3f6b-216">1.5</span></span> |

##### <a name="events"></a><span data-ttu-id="e3f6b-217">イベント</span><span class="sxs-lookup"><span data-stu-id="e3f6b-217">Events</span></span>

<span data-ttu-id="e3f6b-218">[Addハンドラ async](#addhandlerasynceventtype-handler-options-callback)と[removeハンドラ async](#removehandlerasynceventtype-options-callback)を使用して、次のイベントにサブスクライブし、サブスクライブを解除することができます。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-218">You can subscribe to and unsubscribe from the following events using [addHandlerAsync](#addhandlerasynceventtype-handler-options-callback) and [removeHandlerAsync](#removehandlerasynceventtype-options-callback) respectively.</span></span>

| <span data-ttu-id="e3f6b-219">イベント</span><span class="sxs-lookup"><span data-stu-id="e3f6b-219">Event</span></span> | <span data-ttu-id="e3f6b-220">説明</span><span class="sxs-lookup"><span data-stu-id="e3f6b-220">Description</span></span> | <span data-ttu-id="e3f6b-221">最小値</span><span class="sxs-lookup"><span data-stu-id="e3f6b-221">Minimum</span></span><br><span data-ttu-id="e3f6b-222">要件セット</span><span class="sxs-lookup"><span data-stu-id="e3f6b-222">requirement set</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="e3f6b-223">作業ウィンドウが固定されている間、別の Outlook アイテムが選択され、表示することができます。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-223">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="e3f6b-224">1.5</span><span class="sxs-lookup"><span data-stu-id="e3f6b-224">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="e3f6b-225">メールボックスの Office テーマが変更されました。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-225">The Office theme on the mailbox has changed.</span></span> | <span data-ttu-id="e3f6b-226">プレビュー</span><span class="sxs-lookup"><span data-stu-id="e3f6b-226">Preview</span></span> |

### <a name="namespaces"></a><span data-ttu-id="e3f6b-227">名前空間</span><span class="sxs-lookup"><span data-stu-id="e3f6b-227">Namespaces</span></span>

<span data-ttu-id="e3f6b-228">[diagnostics](Office.context.mailbox.diagnostics.md):Outlook アドインに診断情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-228">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="e3f6b-229">[item](Office.context.mailbox.item.md):Outlook アドインのメッセージや予定にアクセスするためのメソッドとプロパティを提供します。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-229">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="e3f6b-230">[userProfile](Office.context.mailbox.userProfile.md):Outlook アドインのユーザーに関する情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-230">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

## <a name="property-details"></a><span data-ttu-id="e3f6b-231">プロパティの詳細</span><span class="sxs-lookup"><span data-stu-id="e3f6b-231">Property details</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="e3f6b-232">ewsUrl: String</span><span class="sxs-lookup"><span data-stu-id="e3f6b-232">ewsUrl: String</span></span>

<span data-ttu-id="e3f6b-p101">このメール アカウントの Exchange Web サービス (EWS) エンドポイントの URL を取得します。読み取りモードのみです。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="e3f6b-235">このメンバーは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-235">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="e3f6b-p102">
  `ewsUrl\` 値は、リモート サービスで、ユーザーのメールボックスに EWS 呼び出しを行うために使うことができます。たとえば、[選択したアイテムから添付ファイルを取得する](/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="e3f6b-238">アプリが閲覧モードで `ewsUrl` メンバーを呼び出すには、アプリのマニフェスト内に **ReadItem** アクセス許可が指定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-238">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="e3f6b-p103">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出してから、`ewsUrl` メンバーを使用する必要があります。アプリには、`saveAsync` メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="e3f6b-241">型</span><span class="sxs-lookup"><span data-stu-id="e3f6b-241">Type</span></span>

*   <span data-ttu-id="e3f6b-242">String</span><span class="sxs-lookup"><span data-stu-id="e3f6b-242">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e3f6b-243">Requirements</span><span class="sxs-lookup"><span data-stu-id="e3f6b-243">Requirements</span></span>

|<span data-ttu-id="e3f6b-244">要件</span><span class="sxs-lookup"><span data-stu-id="e3f6b-244">Requirement</span></span>| <span data-ttu-id="e3f6b-245">値</span><span class="sxs-lookup"><span data-stu-id="e3f6b-245">Value</span></span>|
|---|---|
|[<span data-ttu-id="e3f6b-246">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e3f6b-246">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e3f6b-247">1.0</span><span class="sxs-lookup"><span data-stu-id="e3f6b-247">1.0</span></span>|
|[<span data-ttu-id="e3f6b-248">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e3f6b-248">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e3f6b-249">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e3f6b-249">ReadItem</span></span>|
|[<span data-ttu-id="e3f6b-250">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e3f6b-250">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e3f6b-251">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e3f6b-251">Compose or Read</span></span>|

<br>

---
---

#### <a name="mastercategories-mastercategoriesjavascriptapioutlookofficemastercategories"></a><span data-ttu-id="e3f6b-252">masterCategories: [Mastercategories](/javascript/api/outlook/office.mastercategories)</span><span class="sxs-lookup"><span data-stu-id="e3f6b-252">masterCategories: [MasterCategories](/javascript/api/outlook/office.mastercategories)</span></span>

<span data-ttu-id="e3f6b-253">このメールボックスのカテゴリマスターリストを管理するためのメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-253">Gets an object that provides methods to manage the categories master list on this mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="e3f6b-254">このメンバーは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-254">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="e3f6b-255">種類</span><span class="sxs-lookup"><span data-stu-id="e3f6b-255">Type</span></span>

*   [<span data-ttu-id="e3f6b-256">MasterCategories</span><span class="sxs-lookup"><span data-stu-id="e3f6b-256">MasterCategories</span></span>](/javascript/api/outlook/office.mastercategories)

##### <a name="requirements"></a><span data-ttu-id="e3f6b-257">Requirements</span><span class="sxs-lookup"><span data-stu-id="e3f6b-257">Requirements</span></span>

|<span data-ttu-id="e3f6b-258">要件</span><span class="sxs-lookup"><span data-stu-id="e3f6b-258">Requirement</span></span>| <span data-ttu-id="e3f6b-259">値</span><span class="sxs-lookup"><span data-stu-id="e3f6b-259">Value</span></span>|
|---|---|
|[<span data-ttu-id="e3f6b-260">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e3f6b-260">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e3f6b-261">1.8</span><span class="sxs-lookup"><span data-stu-id="e3f6b-261">1.8</span></span> |
|[<span data-ttu-id="e3f6b-262">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e3f6b-262">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e3f6b-263">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="e3f6b-263">ReadWriteMailbox</span></span> |
|[<span data-ttu-id="e3f6b-264">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e3f6b-264">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e3f6b-265">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e3f6b-265">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="e3f6b-266">例</span><span class="sxs-lookup"><span data-stu-id="e3f6b-266">Example</span></span>

<span data-ttu-id="e3f6b-267">この例では、このメールボックスのカテゴリマスターリストを取得します。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-267">This example gets the categories master list for this mailbox.</span></span>

```js
Office.context.mailbox.masterCategories.getAsync(function (asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    console.log("Action failed with error: " + asyncResult.error.message);
  } else {
    console.log("Master categories: " + JSON.stringify(asyncResult.value));
  }
});
```

<br>

---
---

#### <a name="resturl-string"></a><span data-ttu-id="e3f6b-268">restUrl: String</span><span class="sxs-lookup"><span data-stu-id="e3f6b-268">restUrl: String</span></span>

<span data-ttu-id="e3f6b-269">この電子メール アカウントの REST エンドポイントの URL を取得します。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-269">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="e3f6b-270">`restUrl` 値は、ユーザーのメールボックスに [REST API](/outlook/rest/) 呼び出しを行うために使用できます。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-270">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="e3f6b-271">型</span><span class="sxs-lookup"><span data-stu-id="e3f6b-271">Type</span></span>

*   <span data-ttu-id="e3f6b-272">String</span><span class="sxs-lookup"><span data-stu-id="e3f6b-272">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e3f6b-273">Requirements</span><span class="sxs-lookup"><span data-stu-id="e3f6b-273">Requirements</span></span>

|<span data-ttu-id="e3f6b-274">要件</span><span class="sxs-lookup"><span data-stu-id="e3f6b-274">Requirement</span></span>| <span data-ttu-id="e3f6b-275">値</span><span class="sxs-lookup"><span data-stu-id="e3f6b-275">Value</span></span>|
|---|---|
|[<span data-ttu-id="e3f6b-276">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e3f6b-276">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e3f6b-277">1.5</span><span class="sxs-lookup"><span data-stu-id="e3f6b-277">1.5</span></span> |
|[<span data-ttu-id="e3f6b-278">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e3f6b-278">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e3f6b-279">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e3f6b-279">ReadItem</span></span>|
|[<span data-ttu-id="e3f6b-280">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e3f6b-280">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e3f6b-281">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e3f6b-281">Compose or Read</span></span>|

## <a name="method-details"></a><span data-ttu-id="e3f6b-282">メソッドの詳細</span><span class="sxs-lookup"><span data-stu-id="e3f6b-282">Method details</span></span>

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="e3f6b-283">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="e3f6b-283">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="e3f6b-284">サポートされているイベントのイベント ハンドラーを追加します。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-284">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="e3f6b-285">現在、サポートされている`Office.EventType.ItemChanged`イベント`Office.EventType.OfficeThemeChanged`の種類はおよびです。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-285">Currently, the supported event types are `Office.EventType.ItemChanged` and `Office.EventType.OfficeThemeChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e3f6b-286">パラメーター</span><span class="sxs-lookup"><span data-stu-id="e3f6b-286">Parameters</span></span>

| <span data-ttu-id="e3f6b-287">名前</span><span class="sxs-lookup"><span data-stu-id="e3f6b-287">Name</span></span> | <span data-ttu-id="e3f6b-288">種類</span><span class="sxs-lookup"><span data-stu-id="e3f6b-288">Type</span></span> | <span data-ttu-id="e3f6b-289">属性</span><span class="sxs-lookup"><span data-stu-id="e3f6b-289">Attributes</span></span> | <span data-ttu-id="e3f6b-290">説明</span><span class="sxs-lookup"><span data-stu-id="e3f6b-290">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="e3f6b-291">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="e3f6b-291">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="e3f6b-292">ハンドラーを呼び出す必要のあるイベント。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-292">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="e3f6b-293">Function</span><span class="sxs-lookup"><span data-stu-id="e3f6b-293">Function</span></span> || <span data-ttu-id="e3f6b-p104">イベントを処理する関数。関数は、オブジェクト リテラルである単一パラメーターを受け入れる必要があります。パラメーターの `type` プロパティは、`addHandlerAsync` に渡される `eventType` パラメーターと一致します。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-p104">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="e3f6b-297">Object</span><span class="sxs-lookup"><span data-stu-id="e3f6b-297">Object</span></span> | <span data-ttu-id="e3f6b-298">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="e3f6b-298">&lt;optional&gt;</span></span> | <span data-ttu-id="e3f6b-299">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-299">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="e3f6b-300">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="e3f6b-300">Object</span></span> | <span data-ttu-id="e3f6b-301">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="e3f6b-301">&lt;optional&gt;</span></span> | <span data-ttu-id="e3f6b-302">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-302">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="e3f6b-303">function</span><span class="sxs-lookup"><span data-stu-id="e3f6b-303">function</span></span>| <span data-ttu-id="e3f6b-304">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="e3f6b-304">&lt;optional&gt;</span></span>|<span data-ttu-id="e3f6b-305">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-305">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e3f6b-306">要件</span><span class="sxs-lookup"><span data-stu-id="e3f6b-306">Requirements</span></span>

|<span data-ttu-id="e3f6b-307">要件</span><span class="sxs-lookup"><span data-stu-id="e3f6b-307">Requirement</span></span>| <span data-ttu-id="e3f6b-308">値</span><span class="sxs-lookup"><span data-stu-id="e3f6b-308">Value</span></span>|
|---|---|
|[<span data-ttu-id="e3f6b-309">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e3f6b-309">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e3f6b-310">1.5</span><span class="sxs-lookup"><span data-stu-id="e3f6b-310">1.5</span></span> |
|[<span data-ttu-id="e3f6b-311">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e3f6b-311">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e3f6b-312">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e3f6b-312">ReadItem</span></span> |
|[<span data-ttu-id="e3f6b-313">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e3f6b-313">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e3f6b-314">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e3f6b-314">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e3f6b-315">例</span><span class="sxs-lookup"><span data-stu-id="e3f6b-315">Example</span></span>

```js
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, loadNewItem, function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        // Handle error.
      }
    });
  });
};

function loadNewItem(eventArgs) {
  // Load the properties of the newly selected item.
  loadProps(Office.context.mailbox.item);
}
```

<br>

---
---

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="e3f6b-316">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="e3f6b-316">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="e3f6b-317">REST 形式のアイテム ID を EWS 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-317">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="e3f6b-318">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-318">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="e3f6b-p105">REST API ([Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) や [Microsoft Graph](https://graph.microsoft.io/) など) で取得されたアイテム ID は、Exchange Web サービス (EWS) に使用される形式とは異なる形式を使用します。`convertToEwsId` メソッドは、REST 形式の ID を EWS 用の適切な形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-p105">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e3f6b-321">パラメーター</span><span class="sxs-lookup"><span data-stu-id="e3f6b-321">Parameters</span></span>

|<span data-ttu-id="e3f6b-322">名前</span><span class="sxs-lookup"><span data-stu-id="e3f6b-322">Name</span></span>| <span data-ttu-id="e3f6b-323">型</span><span class="sxs-lookup"><span data-stu-id="e3f6b-323">Type</span></span>| <span data-ttu-id="e3f6b-324">説明</span><span class="sxs-lookup"><span data-stu-id="e3f6b-324">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="e3f6b-325">String</span><span class="sxs-lookup"><span data-stu-id="e3f6b-325">String</span></span>|<span data-ttu-id="e3f6b-326">Outlook REST API 形式のアイテム ID</span><span class="sxs-lookup"><span data-stu-id="e3f6b-326">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="e3f6b-327">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="e3f6b-327">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion)|<span data-ttu-id="e3f6b-328">アイテム ID の取得に使用された Outlook REST API のバージョンを示す値。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-328">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e3f6b-329">Requirements</span><span class="sxs-lookup"><span data-stu-id="e3f6b-329">Requirements</span></span>

|<span data-ttu-id="e3f6b-330">要件</span><span class="sxs-lookup"><span data-stu-id="e3f6b-330">Requirement</span></span>| <span data-ttu-id="e3f6b-331">値</span><span class="sxs-lookup"><span data-stu-id="e3f6b-331">Value</span></span>|
|---|---|
|[<span data-ttu-id="e3f6b-332">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e3f6b-332">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e3f6b-333">1.3</span><span class="sxs-lookup"><span data-stu-id="e3f6b-333">1.3</span></span>|
|[<span data-ttu-id="e3f6b-334">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e3f6b-334">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e3f6b-335">制限あり</span><span class="sxs-lookup"><span data-stu-id="e3f6b-335">Restricted</span></span>|
|[<span data-ttu-id="e3f6b-336">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e3f6b-336">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e3f6b-337">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e3f6b-337">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e3f6b-338">戻り値:</span><span class="sxs-lookup"><span data-stu-id="e3f6b-338">Returns:</span></span>

<span data-ttu-id="e3f6b-339">型:String</span><span class="sxs-lookup"><span data-stu-id="e3f6b-339">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="e3f6b-340">例</span><span class="sxs-lookup"><span data-stu-id="e3f6b-340">Example</span></span>

```js
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime"></a><span data-ttu-id="e3f6b-341">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="e3f6b-341">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime)}</span></span>

<span data-ttu-id="e3f6b-342">クライアントのローカル時間で時間情報が含まれている辞書を取得します。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-342">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="e3f6b-p106">Outlook on the web または Outlook デスクトップのメールアプリでは、日付と時刻に異なるタイムゾーンを使用できます。Outlook デスクトップは、クライアント コンピューターのタイム ゾーンを使用します。Outlook on the web は、Exchange 管理センター (EAC) で設定されたタイム ゾーンを使用します。日付と時刻の値は、ユーザー インターフェイスに表示される値が、常にユーザーが期待するタイム ゾーンと一致するように処理する必要があります。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-p106">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times. Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="e3f6b-p107">Outlook デスクトップ クライアントでメール アプリを実行している場合、`convertToLocalClientTime` メソッドは、クライアント コンピューターのタイム ゾーンに設定された値のディクショナリ オブジェクトを返します。Outlook on the web でメール アプリを実行している場合、`convertToLocalClientTime` メソッドは、EAC で指定したタイム ゾーンに設定された値のディクショナリ オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-p107">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e3f6b-348">パラメーター</span><span class="sxs-lookup"><span data-stu-id="e3f6b-348">Parameters</span></span>

|<span data-ttu-id="e3f6b-349">名前</span><span class="sxs-lookup"><span data-stu-id="e3f6b-349">Name</span></span>| <span data-ttu-id="e3f6b-350">型</span><span class="sxs-lookup"><span data-stu-id="e3f6b-350">Type</span></span>| <span data-ttu-id="e3f6b-351">説明</span><span class="sxs-lookup"><span data-stu-id="e3f6b-351">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="e3f6b-352">日付</span><span class="sxs-lookup"><span data-stu-id="e3f6b-352">Date</span></span>|<span data-ttu-id="e3f6b-353">日付オブジェクト</span><span class="sxs-lookup"><span data-stu-id="e3f6b-353">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e3f6b-354">Requirements</span><span class="sxs-lookup"><span data-stu-id="e3f6b-354">Requirements</span></span>

|<span data-ttu-id="e3f6b-355">要件</span><span class="sxs-lookup"><span data-stu-id="e3f6b-355">Requirement</span></span>| <span data-ttu-id="e3f6b-356">値</span><span class="sxs-lookup"><span data-stu-id="e3f6b-356">Value</span></span>|
|---|---|
|[<span data-ttu-id="e3f6b-357">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e3f6b-357">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e3f6b-358">1.0</span><span class="sxs-lookup"><span data-stu-id="e3f6b-358">1.0</span></span>|
|[<span data-ttu-id="e3f6b-359">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e3f6b-359">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e3f6b-360">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e3f6b-360">ReadItem</span></span>|
|[<span data-ttu-id="e3f6b-361">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e3f6b-361">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e3f6b-362">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e3f6b-362">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e3f6b-363">戻り値:</span><span class="sxs-lookup"><span data-stu-id="e3f6b-363">Returns:</span></span>

<span data-ttu-id="e3f6b-364">型:[LocalClientTime](/javascript/api/outlook/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="e3f6b-364">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime)</span></span>

<br>

---
---

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="e3f6b-365">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="e3f6b-365">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="e3f6b-366">EWS 形式のアイテム ID を REST 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-366">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="e3f6b-367">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-367">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="e3f6b-p108">EWS または `itemId` プロパティで取得されるアイテム ID は、REST API ([Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) や [Microsoft Graph](https://graph.microsoft.io/) など) に使用される形式とは異なる形式を使用します。`convertToRestId` メソッドは、EWS 形式の ID を REST 用の適切な形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-p108">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e3f6b-370">パラメーター</span><span class="sxs-lookup"><span data-stu-id="e3f6b-370">Parameters</span></span>

|<span data-ttu-id="e3f6b-371">名前</span><span class="sxs-lookup"><span data-stu-id="e3f6b-371">Name</span></span>| <span data-ttu-id="e3f6b-372">型</span><span class="sxs-lookup"><span data-stu-id="e3f6b-372">Type</span></span>| <span data-ttu-id="e3f6b-373">説明</span><span class="sxs-lookup"><span data-stu-id="e3f6b-373">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="e3f6b-374">String</span><span class="sxs-lookup"><span data-stu-id="e3f6b-374">String</span></span>|<span data-ttu-id="e3f6b-375">Exchange Web サービス (EWS) 形式のアイテム ID</span><span class="sxs-lookup"><span data-stu-id="e3f6b-375">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="e3f6b-376">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="e3f6b-376">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion)|<span data-ttu-id="e3f6b-377">変換後の ID を使用する Outlook REST API のバージョンを示す値。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-377">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e3f6b-378">Requirements</span><span class="sxs-lookup"><span data-stu-id="e3f6b-378">Requirements</span></span>

|<span data-ttu-id="e3f6b-379">要件</span><span class="sxs-lookup"><span data-stu-id="e3f6b-379">Requirement</span></span>| <span data-ttu-id="e3f6b-380">値</span><span class="sxs-lookup"><span data-stu-id="e3f6b-380">Value</span></span>|
|---|---|
|[<span data-ttu-id="e3f6b-381">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e3f6b-381">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e3f6b-382">1.3</span><span class="sxs-lookup"><span data-stu-id="e3f6b-382">1.3</span></span>|
|[<span data-ttu-id="e3f6b-383">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e3f6b-383">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e3f6b-384">制限あり</span><span class="sxs-lookup"><span data-stu-id="e3f6b-384">Restricted</span></span>|
|[<span data-ttu-id="e3f6b-385">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e3f6b-385">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e3f6b-386">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e3f6b-386">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e3f6b-387">戻り値:</span><span class="sxs-lookup"><span data-stu-id="e3f6b-387">Returns:</span></span>

<span data-ttu-id="e3f6b-388">型:String</span><span class="sxs-lookup"><span data-stu-id="e3f6b-388">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="e3f6b-389">例</span><span class="sxs-lookup"><span data-stu-id="e3f6b-389">Example</span></span>

```js
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="e3f6b-390">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="e3f6b-390">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="e3f6b-391">時間情報が含まれているディクショナリから日付オブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-391">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="e3f6b-392">`convertToUtcClientTime` メソッドは、ローカルの日付と時刻を含むディクショナリを、ローカルの日付と時刻の正しい値を持つ日付オブジェクトに変換します。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-392">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e3f6b-393">パラメーター</span><span class="sxs-lookup"><span data-stu-id="e3f6b-393">Parameters</span></span>

|<span data-ttu-id="e3f6b-394">名前</span><span class="sxs-lookup"><span data-stu-id="e3f6b-394">Name</span></span>| <span data-ttu-id="e3f6b-395">型</span><span class="sxs-lookup"><span data-stu-id="e3f6b-395">Type</span></span>| <span data-ttu-id="e3f6b-396">説明</span><span class="sxs-lookup"><span data-stu-id="e3f6b-396">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="e3f6b-397">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="e3f6b-397">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime)|<span data-ttu-id="e3f6b-398">変換するローカル時刻の値。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-398">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e3f6b-399">Requirements</span><span class="sxs-lookup"><span data-stu-id="e3f6b-399">Requirements</span></span>

|<span data-ttu-id="e3f6b-400">要件</span><span class="sxs-lookup"><span data-stu-id="e3f6b-400">Requirement</span></span>| <span data-ttu-id="e3f6b-401">値</span><span class="sxs-lookup"><span data-stu-id="e3f6b-401">Value</span></span>|
|---|---|
|[<span data-ttu-id="e3f6b-402">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e3f6b-402">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e3f6b-403">1.0</span><span class="sxs-lookup"><span data-stu-id="e3f6b-403">1.0</span></span>|
|[<span data-ttu-id="e3f6b-404">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e3f6b-404">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e3f6b-405">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e3f6b-405">ReadItem</span></span>|
|[<span data-ttu-id="e3f6b-406">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e3f6b-406">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e3f6b-407">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e3f6b-407">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e3f6b-408">戻り値:</span><span class="sxs-lookup"><span data-stu-id="e3f6b-408">Returns:</span></span>

<span data-ttu-id="e3f6b-409">時間が UTC で表現された日付オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-409">A Date object with the time expressed in UTC.</span></span>

<span data-ttu-id="e3f6b-410">型: Date</span><span class="sxs-lookup"><span data-stu-id="e3f6b-410">Type: Date</span></span>

##### <a name="example"></a><span data-ttu-id="e3f6b-411">例</span><span class="sxs-lookup"><span data-stu-id="e3f6b-411">Example</span></span>

```js
// Represents 3:37 PM PDT on Monday, August 26, 2019.
var input = {
  date: 26,
  hours: 15,
  milliseconds: 2,
  minutes: 37,
  month: 7,
  seconds: 2,
  timezoneOffset: -420,
  year: 2019
};

// result should be a Date object.
var result = Office.context.mailbox.convertToUtcClientTime(input);

// Output should be "2019-08-26T22:37:02.002Z".
console.log(result.toISOString());
```

<br>

---
---

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="e3f6b-412">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="e3f6b-412">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="e3f6b-413">既存の予定を表示します。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-413">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="e3f6b-414">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-414">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="e3f6b-415">`displayAppointmentForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存の予定を開きます。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-415">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="e3f6b-p109">Outlook on Mac では、このメソッドを使用して定期的な系列に含まれない単発の予定や定期的な系列のマスター予定を表示できます。ただし、系列のインスタンスは表示できません。これは、Outlook on Mac では定期的な系列のインスタンスのプロパティ (アイテム ID を含む) にアクセスできないためです。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-p109">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="e3f6b-418">Outlook on the web では、このメソッドはフォームの本文が 32KB 以下の文字数の場合にのみ指定のフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-418">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="e3f6b-419">指定のアイテム識別子が既存の予定を表していない場合は、クライアント コンピューターまたはデバイスで空のウィンドウが開き、エラー メッセージは返されません。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-419">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e3f6b-420">パラメーター</span><span class="sxs-lookup"><span data-stu-id="e3f6b-420">Parameters</span></span>

|<span data-ttu-id="e3f6b-421">名前</span><span class="sxs-lookup"><span data-stu-id="e3f6b-421">Name</span></span>| <span data-ttu-id="e3f6b-422">型</span><span class="sxs-lookup"><span data-stu-id="e3f6b-422">Type</span></span>| <span data-ttu-id="e3f6b-423">説明</span><span class="sxs-lookup"><span data-stu-id="e3f6b-423">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="e3f6b-424">String</span><span class="sxs-lookup"><span data-stu-id="e3f6b-424">String</span></span>|<span data-ttu-id="e3f6b-425">既存の予定の Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-425">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e3f6b-426">要件</span><span class="sxs-lookup"><span data-stu-id="e3f6b-426">Requirements</span></span>

|<span data-ttu-id="e3f6b-427">要件</span><span class="sxs-lookup"><span data-stu-id="e3f6b-427">Requirement</span></span>| <span data-ttu-id="e3f6b-428">値</span><span class="sxs-lookup"><span data-stu-id="e3f6b-428">Value</span></span>|
|---|---|
|[<span data-ttu-id="e3f6b-429">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e3f6b-429">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e3f6b-430">1.0</span><span class="sxs-lookup"><span data-stu-id="e3f6b-430">1.0</span></span>|
|[<span data-ttu-id="e3f6b-431">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e3f6b-431">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e3f6b-432">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e3f6b-432">ReadItem</span></span>|
|[<span data-ttu-id="e3f6b-433">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e3f6b-433">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e3f6b-434">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e3f6b-434">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e3f6b-435">例</span><span class="sxs-lookup"><span data-stu-id="e3f6b-435">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

<br>

---
---

#### <a name="displaymessageformitemid"></a><span data-ttu-id="e3f6b-436">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="e3f6b-436">displayMessageForm(itemId)</span></span>

<span data-ttu-id="e3f6b-437">既存のメッセージを表示します。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-437">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="e3f6b-438">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-438">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="e3f6b-439">`displayMessageForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存のメッセージを開きます。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-439">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="e3f6b-440">Outlook on the web では、このメソッドはフォームの本文が 32 KB 以下の文字数の場合にのみ指定のフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-440">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="e3f6b-441">指定のアイテム識別子が既存のメッセージを表していない場合は、クライアント コンピューターにはメッセージは表示されず、エラー メッセージも返されません。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-441">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="e3f6b-p110">予定を表す `itemId` を含む `displayMessageForm` を使用しないでください。既存の予定を表示するには、`displayAppointmentForm` メソッドを使用します。新しい予定を作成するフォームを表示するには、`displayNewAppointmentForm` メソッドを使用します。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-p110">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e3f6b-444">パラメーター</span><span class="sxs-lookup"><span data-stu-id="e3f6b-444">Parameters</span></span>

|<span data-ttu-id="e3f6b-445">名前</span><span class="sxs-lookup"><span data-stu-id="e3f6b-445">Name</span></span>| <span data-ttu-id="e3f6b-446">型</span><span class="sxs-lookup"><span data-stu-id="e3f6b-446">Type</span></span>| <span data-ttu-id="e3f6b-447">説明</span><span class="sxs-lookup"><span data-stu-id="e3f6b-447">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="e3f6b-448">String</span><span class="sxs-lookup"><span data-stu-id="e3f6b-448">String</span></span>|<span data-ttu-id="e3f6b-449">既存のメッセージの Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-449">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e3f6b-450">Requirements</span><span class="sxs-lookup"><span data-stu-id="e3f6b-450">Requirements</span></span>

|<span data-ttu-id="e3f6b-451">要件</span><span class="sxs-lookup"><span data-stu-id="e3f6b-451">Requirement</span></span>| <span data-ttu-id="e3f6b-452">値</span><span class="sxs-lookup"><span data-stu-id="e3f6b-452">Value</span></span>|
|---|---|
|[<span data-ttu-id="e3f6b-453">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e3f6b-453">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e3f6b-454">1.0</span><span class="sxs-lookup"><span data-stu-id="e3f6b-454">1.0</span></span>|
|[<span data-ttu-id="e3f6b-455">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e3f6b-455">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e3f6b-456">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e3f6b-456">ReadItem</span></span>|
|[<span data-ttu-id="e3f6b-457">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e3f6b-457">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e3f6b-458">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e3f6b-458">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e3f6b-459">例</span><span class="sxs-lookup"><span data-stu-id="e3f6b-459">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

<br>

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="e3f6b-460">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="e3f6b-460">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="e3f6b-461">新しい予定を作成するためのフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-461">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="e3f6b-462">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-462">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="e3f6b-p111">`displayNewAppointmentForm` メソッドを使用すると、ユーザーが新しい予定または会議を作成できるフォームが開きます。パラメーターを指定すると、予定のフォーム フィールドにパラメーターの内容が自動的に設定されます。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-p111">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="e3f6b-p112">Outlook on the web およびモバイル デバイスでは、このメソッドは常に出席者フィールドが含まれるフォームを表示します。入力引数として出席者を指定しないと、このメソッドは **[保存]** ボタンのあるフォームを表示します。出席者を指定した場合には、フォームにその出席者と **[送信]** ボタンが表示されます。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-p112">In Outlook on the web and mobile devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="e3f6b-p113">Outlook リッチ クライアントと Outlook RT で、`requiredAttendees`、`optionalAttendees`、または `resources` パラメーターに出席者またはリソースを指定し、このメソッドを実行すると、**[送信]** ボタンがある会議フォームが表示されます。受信者を指定せずにこのメソッドを実行すると、**[保存して閉じる]** ボタンがある予定フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-p113">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="e3f6b-470">パラメーターのいずれかが指定のサイズ制限を超える場合、または不明なパラメーター名が指定されている場合は、例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-470">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e3f6b-471">パラメーター</span><span class="sxs-lookup"><span data-stu-id="e3f6b-471">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="e3f6b-472">すべてのパラメーターは省略可能です。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-472">All parameters are optional.</span></span>

|<span data-ttu-id="e3f6b-473">名前</span><span class="sxs-lookup"><span data-stu-id="e3f6b-473">Name</span></span>| <span data-ttu-id="e3f6b-474">型</span><span class="sxs-lookup"><span data-stu-id="e3f6b-474">Type</span></span>| <span data-ttu-id="e3f6b-475">説明</span><span class="sxs-lookup"><span data-stu-id="e3f6b-475">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="e3f6b-476">Object</span><span class="sxs-lookup"><span data-stu-id="e3f6b-476">Object</span></span> | <span data-ttu-id="e3f6b-477">新しい予定を記述するパラメーターのディクショナリ。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-477">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="e3f6b-478">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="e3f6b-478">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="e3f6b-p114">予定に必要な各出席者について、メール アドレスを含む文字列の配列、または `EmailAddressDetails` オブジェクトを含む配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-p114">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="e3f6b-481">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="e3f6b-481">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="e3f6b-p115">予定の各任意出席者について、メール アドレスを含む文字列の配列、または `EmailAddressDetails` オブジェクトを含む配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-p115">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="e3f6b-484">日付</span><span class="sxs-lookup"><span data-stu-id="e3f6b-484">Date</span></span> | <span data-ttu-id="e3f6b-485">予定の開始日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-485">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="e3f6b-486">日付</span><span class="sxs-lookup"><span data-stu-id="e3f6b-486">Date</span></span> | <span data-ttu-id="e3f6b-487">予定の終了日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-487">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="e3f6b-488">String</span><span class="sxs-lookup"><span data-stu-id="e3f6b-488">String</span></span> | <span data-ttu-id="e3f6b-p116">予定の場所を含む文字列。文字列は最大 255 文字に制限されます。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-p116">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="e3f6b-491">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="e3f6b-491">Array.&lt;String&gt;</span></span> | <span data-ttu-id="e3f6b-p117">予定に必要なリソースを含む文字列の配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-p117">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="e3f6b-494">String</span><span class="sxs-lookup"><span data-stu-id="e3f6b-494">String</span></span> | <span data-ttu-id="e3f6b-p118">予定の件名を含む文字列です。文字列は最大 255 文字に制限されます。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-p118">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="e3f6b-497">String</span><span class="sxs-lookup"><span data-stu-id="e3f6b-497">String</span></span> | <span data-ttu-id="e3f6b-p119">予定の本文。本文の内容は、最大サイズが 32 KB に制限されます。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-p119">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="e3f6b-500">要件</span><span class="sxs-lookup"><span data-stu-id="e3f6b-500">Requirements</span></span>

|<span data-ttu-id="e3f6b-501">要件</span><span class="sxs-lookup"><span data-stu-id="e3f6b-501">Requirement</span></span>| <span data-ttu-id="e3f6b-502">値</span><span class="sxs-lookup"><span data-stu-id="e3f6b-502">Value</span></span>|
|---|---|
|[<span data-ttu-id="e3f6b-503">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e3f6b-503">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e3f6b-504">1.0</span><span class="sxs-lookup"><span data-stu-id="e3f6b-504">1.0</span></span>|
|[<span data-ttu-id="e3f6b-505">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e3f6b-505">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e3f6b-506">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e3f6b-506">ReadItem</span></span>|
|[<span data-ttu-id="e3f6b-507">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e3f6b-507">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e3f6b-508">読み取り</span><span class="sxs-lookup"><span data-stu-id="e3f6b-508">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e3f6b-509">例</span><span class="sxs-lookup"><span data-stu-id="e3f6b-509">Example</span></span>

```js
var start = new Date();
var end = new Date();
end.setHours(start.getHours() + 1);

Office.context.mailbox.displayNewAppointmentForm(
  {
    requiredAttendees: ['bob@contoso.com'],
    optionalAttendees: ['sam@contoso.com'],
    start: start,
    end: end,
    location: 'Home',
    resources: ['projector@contoso.com'],
    subject: 'meeting',
    body: 'Hello World!'
  });
```

<br>

---
---

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="e3f6b-510">displayNewMessageForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="e3f6b-510">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="e3f6b-511">新しいメッセージを作成するためのフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-511">Displays a form for creating a new message.</span></span>

<span data-ttu-id="e3f6b-p120">The `displayNewMessageForm` method opens a form that enables the user to create a new message. If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span><span class="sxs-lookup"><span data-stu-id="e3f6b-p120">The `displayNewMessageForm` method opens a form that enables the user to create a new message. If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="e3f6b-514">パラメーターのいずれかが指定のサイズ制限を超える場合、または不明なパラメーター名が指定されている場合は、例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-514">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e3f6b-515">パラメーター</span><span class="sxs-lookup"><span data-stu-id="e3f6b-515">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="e3f6b-516">すべてのパラメーターは省略可能です。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-516">All parameters are optional.</span></span>

|<span data-ttu-id="e3f6b-517">名前</span><span class="sxs-lookup"><span data-stu-id="e3f6b-517">Name</span></span>| <span data-ttu-id="e3f6b-518">型</span><span class="sxs-lookup"><span data-stu-id="e3f6b-518">Type</span></span>| <span data-ttu-id="e3f6b-519">説明</span><span class="sxs-lookup"><span data-stu-id="e3f6b-519">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="e3f6b-520">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="e3f6b-520">Object</span></span> | <span data-ttu-id="e3f6b-521">新しいメッセージを記述するパラメータの辞書。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-521">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="e3f6b-522">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="e3f6b-522">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="e3f6b-p121">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line. The array is limited to a maximum of 100 entries.</span><span class="sxs-lookup"><span data-stu-id="e3f6b-p121">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="e3f6b-525">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="e3f6b-525">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="e3f6b-p122">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line. The array is limited to a maximum of 100 entries.</span><span class="sxs-lookup"><span data-stu-id="e3f6b-p122">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="e3f6b-528">配列。&lt;文字列&gt; | 配列。&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="e3f6b-528">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="e3f6b-p123">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line. The array is limited to a maximum of 100 entries.</span><span class="sxs-lookup"><span data-stu-id="e3f6b-p123">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="e3f6b-531">String</span><span class="sxs-lookup"><span data-stu-id="e3f6b-531">String</span></span> | <span data-ttu-id="e3f6b-p124">A string containing the subject of the message. The string is limited to a maximum of 255 characters.</span><span class="sxs-lookup"><span data-stu-id="e3f6b-p124">A string containing the subject of the message. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="e3f6b-534">String</span><span class="sxs-lookup"><span data-stu-id="e3f6b-534">String</span></span> | <span data-ttu-id="e3f6b-p125">The HTML body of the message. The body content is limited to a maximum size of 32 KB.</span><span class="sxs-lookup"><span data-stu-id="e3f6b-p125">The HTML body of the message. The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="e3f6b-537">配列。&lt;オブジェクト&gt;</span><span class="sxs-lookup"><span data-stu-id="e3f6b-537">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="e3f6b-538">添付ファイルまたは添付アイテムである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-538">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="e3f6b-539">String</span><span class="sxs-lookup"><span data-stu-id="e3f6b-539">String</span></span> | <span data-ttu-id="e3f6b-p126">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-p126">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="e3f6b-542">String</span><span class="sxs-lookup"><span data-stu-id="e3f6b-542">String</span></span> | <span data-ttu-id="e3f6b-543">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-543">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="e3f6b-544">文字列</span><span class="sxs-lookup"><span data-stu-id="e3f6b-544">String</span></span> | <span data-ttu-id="e3f6b-p127">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-p127">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="e3f6b-547">ブール値</span><span class="sxs-lookup"><span data-stu-id="e3f6b-547">Boolean</span></span> | <span data-ttu-id="e3f6b-p128">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-p128">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="e3f6b-550">String</span><span class="sxs-lookup"><span data-stu-id="e3f6b-550">String</span></span> | <span data-ttu-id="e3f6b-p129">Only used if `type` is set to `item`. The EWS item id of the existing e-mail you want to attach to the new message. This is a string up to 100 characters.</span><span class="sxs-lookup"><span data-stu-id="e3f6b-p129">Only used if `type` is set to `item`. The EWS item id of the existing e-mail you want to attach to the new message. This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="e3f6b-554">Requirements</span><span class="sxs-lookup"><span data-stu-id="e3f6b-554">Requirements</span></span>

|<span data-ttu-id="e3f6b-555">要件</span><span class="sxs-lookup"><span data-stu-id="e3f6b-555">Requirement</span></span>| <span data-ttu-id="e3f6b-556">値</span><span class="sxs-lookup"><span data-stu-id="e3f6b-556">Value</span></span>|
|---|---|
|[<span data-ttu-id="e3f6b-557">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e3f6b-557">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e3f6b-558">1.6</span><span class="sxs-lookup"><span data-stu-id="e3f6b-558">1.6</span></span> |
|[<span data-ttu-id="e3f6b-559">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e3f6b-559">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e3f6b-560">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e3f6b-560">ReadItem</span></span>|
|[<span data-ttu-id="e3f6b-561">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e3f6b-561">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e3f6b-562">読み取り</span><span class="sxs-lookup"><span data-stu-id="e3f6b-562">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e3f6b-563">例</span><span class="sxs-lookup"><span data-stu-id="e3f6b-563">Example</span></span>

```js
Office.context.mailbox.displayNewMessageForm(
  {
    // Copy the To line from current item.
    toRecipients: Office.context.mailbox.item.to,
    ccRecipients: ['sam@contoso.com'],
    subject: 'Outlook add-ins are cool!',
    htmlBody: 'Hello <b>World</b>!<br/><img src="cid:image.png"></i>',
    attachments: [
      {
        type: 'file',
        name: 'image.png',
        url: 'http://contoso.com/image.png',
        isInline: true
      }
    ]
  });
```

<br>

---
---

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="e3f6b-564">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="e3f6b-564">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="e3f6b-565">REST API または Exchange Web サービスを呼び出すために使用するトークンを含む文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-565">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="e3f6b-p130">`getCallbackTokenAsync` メソッドは、ユーザーのメールボックスをホストする Exchange Server から不透明なトークンを取得する非同期の呼び出しを行います。コールバック トークンの有効期間は 5 分です。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-p130">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="e3f6b-568">可能であれば、アドインでは Exchange Web サービスの代わりに REST API を使用することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-568">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span>

<span data-ttu-id="e3f6b-569">閲覧モードで `getCallbackTokenAsync` メソッドを呼び出すには、**ReadItem** の最小限のアクセス許可レベルが必要です。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-569">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="e3f6b-570">作成モードで `getCallbackTokenAsync` を呼び出すには、アイテムを保存しておく必要があります。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-570">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="e3f6b-571">[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドは、**ReadWriteItem** の最小限のアクセス許可レベルが必要です。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-571">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

<span data-ttu-id="e3f6b-572">**REST トークン**</span><span class="sxs-lookup"><span data-stu-id="e3f6b-572">**REST Tokens**</span></span>

<span data-ttu-id="e3f6b-p132">REST トークンが要求された場合 (`options.isRest = true`)、結果トークンは Exchange Web サービスの呼び出しを認証するためには機能しません。アドインがマニフェストで [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) アクセス許可を指定していない限り、トークンの範囲は現在のアイテムとその添付ファイルへの読み取り専用アクセスに制限されます。`ReadWriteMailbox` アクセス許可が指定されている場合は、結果トークンは、メールを送信する機能など、メール、カレンダー、連絡先への読み取り/書き込みアクセスを付与します。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-p132">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="e3f6b-576">アドインでは、`restUrl` プロパティを使用して、REST API 呼び出しを行うときに使用する正しい URL を決定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-576">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="e3f6b-577">**EWS トークン**</span><span class="sxs-lookup"><span data-stu-id="e3f6b-577">**EWS Tokens**</span></span>

<span data-ttu-id="e3f6b-p133">EWS トークンが要求された場合 (`options.isRest = false`)、結果トークンは REST API 呼び出しを認証するためには機能しません。トークンの範囲は、現在のアイテムへのアクセスに制限されます。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-p133">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="e3f6b-580">アドインでは、`ewsUrl` プロパティを使用して、EWS 呼び出しを行うときに使用する正しい URL を決定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-580">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

<span data-ttu-id="e3f6b-581">トークンと、添付ファイル識別子またはアイテム識別子の両方をサードパーティ システムに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-581">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="e3f6b-582">サードパーティシステムは、トークンをベアラー認証トークンとして使用して、Exchange Web サービス (EWS) の[Getattachment](/exchange/client-developer/web-service-reference/getattachment-operation)操作または[GetItem](/exchange/client-developer/web-service-reference/getitem-operation)操作を呼び出して、添付ファイルまたはアイテムを取得します。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-582">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to retrieve an attachment or item.</span></span> <span data-ttu-id="e3f6b-583">たとえば、[選択したアイテムから添付ファイルを取得する](/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-583">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

##### <a name="parameters"></a><span data-ttu-id="e3f6b-584">パラメーター</span><span class="sxs-lookup"><span data-stu-id="e3f6b-584">Parameters</span></span>

|<span data-ttu-id="e3f6b-585">名前</span><span class="sxs-lookup"><span data-stu-id="e3f6b-585">Name</span></span>| <span data-ttu-id="e3f6b-586">型</span><span class="sxs-lookup"><span data-stu-id="e3f6b-586">Type</span></span>| <span data-ttu-id="e3f6b-587">属性</span><span class="sxs-lookup"><span data-stu-id="e3f6b-587">Attributes</span></span>| <span data-ttu-id="e3f6b-588">説明</span><span class="sxs-lookup"><span data-stu-id="e3f6b-588">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="e3f6b-589">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="e3f6b-589">Object</span></span> | <span data-ttu-id="e3f6b-590">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="e3f6b-590">&lt;optional&gt;</span></span> | <span data-ttu-id="e3f6b-591">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-591">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="e3f6b-592">ブール値</span><span class="sxs-lookup"><span data-stu-id="e3f6b-592">Boolean</span></span> |  <span data-ttu-id="e3f6b-593">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="e3f6b-593">&lt;optional&gt;</span></span> | <span data-ttu-id="e3f6b-p135">提供されたトークンを Outlook REST API または Exchange Web サービスに使用するかどうかを決定します。既定値は、`false` です。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-p135">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="e3f6b-596">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="e3f6b-596">Object</span></span> |  <span data-ttu-id="e3f6b-597">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="e3f6b-597">&lt;optional&gt;</span></span> | <span data-ttu-id="e3f6b-598">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-598">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="e3f6b-599">function</span><span class="sxs-lookup"><span data-stu-id="e3f6b-599">function</span></span>||<span data-ttu-id="e3f6b-600">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-600">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="e3f6b-601">トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-601">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="e3f6b-602">エラーが発生した場合、 `asyncResult.error` および `asyncResult.diagnostics` のプロパティで追加情報が提供される場合があります。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-602">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|

##### <a name="errors"></a><span data-ttu-id="e3f6b-603">エラー</span><span class="sxs-lookup"><span data-stu-id="e3f6b-603">Errors</span></span>

|<span data-ttu-id="e3f6b-604">エラー コード</span><span class="sxs-lookup"><span data-stu-id="e3f6b-604">Error code</span></span>|<span data-ttu-id="e3f6b-605">説明</span><span class="sxs-lookup"><span data-stu-id="e3f6b-605">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="e3f6b-606">要求が失敗しました。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-606">The request has failed.</span></span> <span data-ttu-id="e3f6b-607">HTTP エラーコードの diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-607">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="e3f6b-608">Exchange サーバーがエラーを返しました。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-608">The Exchange server returned an error.</span></span> <span data-ttu-id="e3f6b-609">詳細については、diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-609">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="e3f6b-610">ユーザーはネットワークに接続されていません。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-610">The user is no longer connected to the network.</span></span> <span data-ttu-id="e3f6b-611">ネットワーク接続を確認し、やり直してください。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-611">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e3f6b-612">要件</span><span class="sxs-lookup"><span data-stu-id="e3f6b-612">Requirements</span></span>

|<span data-ttu-id="e3f6b-613">必要条件</span><span class="sxs-lookup"><span data-stu-id="e3f6b-613">Requirement</span></span>| <span data-ttu-id="e3f6b-614">値</span><span class="sxs-lookup"><span data-stu-id="e3f6b-614">Value</span></span>|
|---|---|
|[<span data-ttu-id="e3f6b-615">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e3f6b-615">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e3f6b-616">1.5</span><span class="sxs-lookup"><span data-stu-id="e3f6b-616">1.5</span></span> |
|[<span data-ttu-id="e3f6b-617">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e3f6b-617">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e3f6b-618">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e3f6b-618">ReadItem</span></span>|
|[<span data-ttu-id="e3f6b-619">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e3f6b-619">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e3f6b-620">新規作成と閲覧</span><span class="sxs-lookup"><span data-stu-id="e3f6b-620">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="e3f6b-621">例</span><span class="sxs-lookup"><span data-stu-id="e3f6b-621">Example</span></span>

```js
function getCallbackToken() {
  var options = {
    isRest: true,
    asyncContext: { message: 'Hello World!' }
  };

  Office.context.mailbox.getCallbackTokenAsync(options, cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

<br>

---
---

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="e3f6b-622">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="e3f6b-622">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="e3f6b-623">Exchange Server から添付ファイルやアイテムを取得するために使うトークンを含む文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-623">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="e3f6b-p139">`getCallbackTokenAsync` メソッドは、ユーザーのメールボックスをホストする Exchange Server から不透明なトークンを取得する非同期の呼び出しを行います。コールバック トークンの有効期間は 5 分です。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-p139">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="e3f6b-626">トークンと、添付ファイル識別子またはアイテム識別子の両方をサードパーティ システムに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-626">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="e3f6b-627">サードパーティ システムは、トークンをベアラー承認トークンとして使用し、Exchange Web サービス (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) 操作または [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) 操作を呼び出して、添付ファイルまたはアイテムを返します。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-627">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item.</span></span> <span data-ttu-id="e3f6b-628">たとえば、[選択したアイテムから添付ファイルを取得する](/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-628">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="e3f6b-629">閲覧モードで `getCallbackTokenAsync` メソッドを呼び出すには、**ReadItem** の最小限のアクセス許可レベルが必要です。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-629">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="e3f6b-630">作成モードで `getCallbackTokenAsync` を呼び出すには、アイテムを保存しておく必要があります。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-630">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="e3f6b-631">[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドは、**ReadWriteItem** の最小限のアクセス許可レベルが必要です。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-631">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e3f6b-632">パラメーター</span><span class="sxs-lookup"><span data-stu-id="e3f6b-632">Parameters</span></span>

|<span data-ttu-id="e3f6b-633">名前</span><span class="sxs-lookup"><span data-stu-id="e3f6b-633">Name</span></span>| <span data-ttu-id="e3f6b-634">型</span><span class="sxs-lookup"><span data-stu-id="e3f6b-634">Type</span></span>| <span data-ttu-id="e3f6b-635">属性</span><span class="sxs-lookup"><span data-stu-id="e3f6b-635">Attributes</span></span>| <span data-ttu-id="e3f6b-636">説明</span><span class="sxs-lookup"><span data-stu-id="e3f6b-636">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="e3f6b-637">関数</span><span class="sxs-lookup"><span data-stu-id="e3f6b-637">function</span></span>||<span data-ttu-id="e3f6b-638">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-638">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="e3f6b-639">トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-639">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="e3f6b-640">エラーが発生した場合、 `asyncResult.error` および `asyncResult.diagnostics` のプロパティで追加情報が提供される場合があります。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-640">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="e3f6b-641">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="e3f6b-641">Object</span></span>| <span data-ttu-id="e3f6b-642">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="e3f6b-642">&lt;optional&gt;</span></span>|<span data-ttu-id="e3f6b-643">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-643">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="e3f6b-644">エラー</span><span class="sxs-lookup"><span data-stu-id="e3f6b-644">Errors</span></span>

|<span data-ttu-id="e3f6b-645">エラー コード</span><span class="sxs-lookup"><span data-stu-id="e3f6b-645">Error code</span></span>|<span data-ttu-id="e3f6b-646">説明</span><span class="sxs-lookup"><span data-stu-id="e3f6b-646">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="e3f6b-647">要求が失敗しました。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-647">The request has failed.</span></span> <span data-ttu-id="e3f6b-648">HTTP エラーコードの diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-648">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="e3f6b-649">Exchange サーバーがエラーを返しました。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-649">The Exchange server returned an error.</span></span> <span data-ttu-id="e3f6b-650">詳細については、diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-650">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="e3f6b-651">ユーザーはネットワークに接続されていません。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-651">The user is no longer connected to the network.</span></span> <span data-ttu-id="e3f6b-652">ネットワーク接続を確認し、やり直してください。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-652">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e3f6b-653">要件</span><span class="sxs-lookup"><span data-stu-id="e3f6b-653">Requirements</span></span>

|<span data-ttu-id="e3f6b-654">要件</span><span class="sxs-lookup"><span data-stu-id="e3f6b-654">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="e3f6b-655">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e3f6b-655">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e3f6b-656">1.0</span><span class="sxs-lookup"><span data-stu-id="e3f6b-656">1.0</span></span> | <span data-ttu-id="e3f6b-657">1.3</span><span class="sxs-lookup"><span data-stu-id="e3f6b-657">1.3</span></span> |
|[<span data-ttu-id="e3f6b-658">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e3f6b-658">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e3f6b-659">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e3f6b-659">ReadItem</span></span> | <span data-ttu-id="e3f6b-660">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e3f6b-660">ReadItem</span></span> |
|[<span data-ttu-id="e3f6b-661">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e3f6b-661">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e3f6b-662">Read</span><span class="sxs-lookup"><span data-stu-id="e3f6b-662">Read</span></span> | <span data-ttu-id="e3f6b-663">Compose</span><span class="sxs-lookup"><span data-stu-id="e3f6b-663">Compose</span></span> |

##### <a name="example"></a><span data-ttu-id="e3f6b-664">例</span><span class="sxs-lookup"><span data-stu-id="e3f6b-664">Example</span></span>

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

<br>

---
---

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="e3f6b-665">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="e3f6b-665">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="e3f6b-666">ユーザーと Office アドインを識別するトークンを取得します。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-666">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="e3f6b-667">`getUserIdentityTokenAsync` メソッドは、[アドインとユーザーをサード パーティのシステムで識別して認証](/outlook/add-ins/authentication)することのできるトークンを返します。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-667">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="e3f6b-668">パラメーター</span><span class="sxs-lookup"><span data-stu-id="e3f6b-668">Parameters</span></span>

|<span data-ttu-id="e3f6b-669">名前</span><span class="sxs-lookup"><span data-stu-id="e3f6b-669">Name</span></span>| <span data-ttu-id="e3f6b-670">型</span><span class="sxs-lookup"><span data-stu-id="e3f6b-670">Type</span></span>| <span data-ttu-id="e3f6b-671">属性</span><span class="sxs-lookup"><span data-stu-id="e3f6b-671">Attributes</span></span>| <span data-ttu-id="e3f6b-672">説明</span><span class="sxs-lookup"><span data-stu-id="e3f6b-672">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="e3f6b-673">関数</span><span class="sxs-lookup"><span data-stu-id="e3f6b-673">function</span></span>||<span data-ttu-id="e3f6b-674">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-674">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="e3f6b-675">トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-675">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="e3f6b-676">エラーが発生した場合、 `asyncResult.error` および `asyncResult.diagnostics` のプロパティで追加情報が提供される場合があります。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-676">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="e3f6b-677">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="e3f6b-677">Object</span></span>| <span data-ttu-id="e3f6b-678">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="e3f6b-678">&lt;optional&gt;</span></span>|<span data-ttu-id="e3f6b-679">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-679">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="e3f6b-680">エラー</span><span class="sxs-lookup"><span data-stu-id="e3f6b-680">Errors</span></span>

|<span data-ttu-id="e3f6b-681">エラー コード</span><span class="sxs-lookup"><span data-stu-id="e3f6b-681">Error code</span></span>|<span data-ttu-id="e3f6b-682">説明</span><span class="sxs-lookup"><span data-stu-id="e3f6b-682">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="e3f6b-683">要求が失敗しました。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-683">The request has failed.</span></span> <span data-ttu-id="e3f6b-684">HTTP エラーコードの diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-684">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="e3f6b-685">Exchange サーバーがエラーを返しました。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-685">The Exchange server returned an error.</span></span> <span data-ttu-id="e3f6b-686">詳細については、diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-686">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="e3f6b-687">ユーザーはネットワークに接続されていません。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-687">The user is no longer connected to the network.</span></span> <span data-ttu-id="e3f6b-688">ネットワーク接続を確認し、やり直してください。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-688">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e3f6b-689">要件</span><span class="sxs-lookup"><span data-stu-id="e3f6b-689">Requirements</span></span>

|<span data-ttu-id="e3f6b-690">要件</span><span class="sxs-lookup"><span data-stu-id="e3f6b-690">Requirement</span></span>| <span data-ttu-id="e3f6b-691">値</span><span class="sxs-lookup"><span data-stu-id="e3f6b-691">Value</span></span>|
|---|---|
|[<span data-ttu-id="e3f6b-692">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e3f6b-692">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e3f6b-693">1.0</span><span class="sxs-lookup"><span data-stu-id="e3f6b-693">1.0</span></span>|
|[<span data-ttu-id="e3f6b-694">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e3f6b-694">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e3f6b-695">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e3f6b-695">ReadItem</span></span>|
|[<span data-ttu-id="e3f6b-696">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e3f6b-696">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e3f6b-697">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e3f6b-697">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e3f6b-698">例</span><span class="sxs-lookup"><span data-stu-id="e3f6b-698">Example</span></span>

```js
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

<br>

---
---

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="e3f6b-699">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="e3f6b-699">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="e3f6b-700">ユーザーのメールボックスをホストしている Exchange サーバー上の Exchange Web サービス (EWS) のサービスに対して非同期の要求を行います。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-700">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="e3f6b-701">このメソッドは、次のシナリオではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-701">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="e3f6b-702">Outlook on iOS または Android の場合</span><span class="sxs-lookup"><span data-stu-id="e3f6b-702">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="e3f6b-703">アドインが Gmail のメールボックスに読み込まれる場合</span><span class="sxs-lookup"><span data-stu-id="e3f6b-703">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="e3f6b-704">このような場合は、アドインでは [REST API を使用](/outlook/add-ins/use-rest-api)して、代わりにユーザーのメールボックスにアクセスする必要があります。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-704">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="e3f6b-p148">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange. See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span><span class="sxs-lookup"><span data-stu-id="e3f6b-p148">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange. See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="e3f6b-707">`makeEwsRequestAsync` メソッドでは、フォルダー関連アイテムを要求できません。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-707">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="e3f6b-708">XML 要求では UTF-8 エンコードを指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-708">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="e3f6b-p149">`makeEwsRequestAsync` メソッドを使用するには、アドインに **ReadWriteMailbox** アクセス許可が必要です。**ReadWriteMailbox** アクセス許可と、`makeEwsRequestAsync` メソッドで呼び出せる EWS 操作の使い方については、「[ユーザーのメールボックスへのメール アドイン アクセスのアクセス許可を指定する](/outlook/add-ins/understanding-outlook-add-in-permissions)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-p149">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="e3f6b-711">サーバー管理者は、クライアント アクセス サーバーの EWS ディレクトリで `OAuthAuthentication` を true に設定して、`makeEwsRequestAsync` メソッドで EWS 要求を行うことができるようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-711">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="e3f6b-712">バージョンの相違点</span><span class="sxs-lookup"><span data-stu-id="e3f6b-712">Version differences</span></span>

<span data-ttu-id="e3f6b-713">バージョン 15.0.4535.1004 より前のバージョンの Outlook で実行しているメール アプリで `makeEwsRequestAsync` メソッドを使用する場合には、エンコード値を `ISO-8859-1` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-713">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="e3f6b-714">Outlook on the web でメール アプリを実行している場合は、エンコード値を設定する必要はありません。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-714">You do not need to set the encoding value when your mail app is running in Outlook on the web.</span></span> <span data-ttu-id="e3f6b-715">メールアプリが web 上の Outlook またはデスクトップクライアントで実行されているかどうかは、mailbox プロパティを使用して判断できます。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-715">You can determine whether your mail app is running in Outlook on the web or a desktop client by using the mailbox.diagnostics.hostName property.</span></span> <span data-ttu-id="e3f6b-716">mailbox.diagnostics.hostVersion プロパティを使って、どのバージョンの Outlook を使って実行しているかを確認できます。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-716">You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e3f6b-717">パラメーター</span><span class="sxs-lookup"><span data-stu-id="e3f6b-717">Parameters</span></span>

|<span data-ttu-id="e3f6b-718">名前</span><span class="sxs-lookup"><span data-stu-id="e3f6b-718">Name</span></span>| <span data-ttu-id="e3f6b-719">型</span><span class="sxs-lookup"><span data-stu-id="e3f6b-719">Type</span></span>| <span data-ttu-id="e3f6b-720">属性</span><span class="sxs-lookup"><span data-stu-id="e3f6b-720">Attributes</span></span>| <span data-ttu-id="e3f6b-721">説明</span><span class="sxs-lookup"><span data-stu-id="e3f6b-721">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="e3f6b-722">String</span><span class="sxs-lookup"><span data-stu-id="e3f6b-722">String</span></span>||<span data-ttu-id="e3f6b-723">EWS 要求です。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-723">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="e3f6b-724">function</span><span class="sxs-lookup"><span data-stu-id="e3f6b-724">function</span></span>||<span data-ttu-id="e3f6b-725">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-725">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="e3f6b-p151">The XML result of the EWS call is provided as a string in the `asyncResult.value` property. If the result exceeds 1 MB in size, an error message is returned instead.</span><span class="sxs-lookup"><span data-stu-id="e3f6b-p151">The XML result of the EWS call is provided as a string in the `asyncResult.value` property. If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="e3f6b-728">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="e3f6b-728">Object</span></span>| <span data-ttu-id="e3f6b-729">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="e3f6b-729">&lt;optional&gt;</span></span>|<span data-ttu-id="e3f6b-730">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-730">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e3f6b-731">要件</span><span class="sxs-lookup"><span data-stu-id="e3f6b-731">Requirements</span></span>

|<span data-ttu-id="e3f6b-732">要件</span><span class="sxs-lookup"><span data-stu-id="e3f6b-732">Requirement</span></span>| <span data-ttu-id="e3f6b-733">値</span><span class="sxs-lookup"><span data-stu-id="e3f6b-733">Value</span></span>|
|---|---|
|[<span data-ttu-id="e3f6b-734">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e3f6b-734">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e3f6b-735">1.0</span><span class="sxs-lookup"><span data-stu-id="e3f6b-735">1.0</span></span>|
|[<span data-ttu-id="e3f6b-736">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e3f6b-736">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e3f6b-737">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="e3f6b-737">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="e3f6b-738">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e3f6b-738">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e3f6b-739">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e3f6b-739">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e3f6b-740">例</span><span class="sxs-lookup"><span data-stu-id="e3f6b-740">Example</span></span>

<span data-ttu-id="e3f6b-741">次の例は、`GetItem` 操作を使ってアイテムの件名を取得するため、`makeEwsRequestAsync` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-741">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

```js
function getSubjectRequest(id) {
  // Return a GetItem operation request for the subject of the specified item.
  var request =
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
    '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
    '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
    '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
    '  <soap:Header>' +
    '    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
    '  </soap:Header>' +
    '  <soap:Body>' +
    '    <GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
    '      <ItemShape>' +
    '        <t:BaseShape>IdOnly</t:BaseShape>' +
    '        <t:AdditionalProperties>' +
    '            <t:FieldURI FieldURI="item:Subject"/>' +
    '        </t:AdditionalProperties>' +
    '      </ItemShape>' +
    '      <ItemIds><t:ItemId Id="' + id + '"/></ItemIds>' +
    '    </GetItem>' +
    '  </soap:Body>' +
    '</soap:Envelope>';

  return request;
}

function sendRequest() {
  // Create a local variable that contains the mailbox.
  Office.context.mailbox.makeEwsRequestAsync(
    getSubjectRequest(mailbox.item.itemId), callback);
}

function callback(asyncResult)  {
  var result = asyncResult.value;
  var context = asyncResult.asyncContext;

  // Process the returned response here.
}
```

<br>

---
---

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="e3f6b-742">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="e3f6b-742">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="e3f6b-743">サポートされているイベントの種類のイベント ハンドラーを削除します。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-743">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="e3f6b-744">現在、サポートされている`Office.EventType.ItemChanged`イベント`Office.EventType.OfficeThemeChanged`の種類はおよびです。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-744">Currently, the supported event types are `Office.EventType.ItemChanged` and `Office.EventType.OfficeThemeChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e3f6b-745">パラメーター</span><span class="sxs-lookup"><span data-stu-id="e3f6b-745">Parameters</span></span>

| <span data-ttu-id="e3f6b-746">名前</span><span class="sxs-lookup"><span data-stu-id="e3f6b-746">Name</span></span> | <span data-ttu-id="e3f6b-747">型</span><span class="sxs-lookup"><span data-stu-id="e3f6b-747">Type</span></span> | <span data-ttu-id="e3f6b-748">属性</span><span class="sxs-lookup"><span data-stu-id="e3f6b-748">Attributes</span></span> | <span data-ttu-id="e3f6b-749">説明</span><span class="sxs-lookup"><span data-stu-id="e3f6b-749">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="e3f6b-750">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="e3f6b-750">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="e3f6b-751">ハンドラーを取り消すイベント。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-751">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="e3f6b-752">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="e3f6b-752">Object</span></span> | <span data-ttu-id="e3f6b-753">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="e3f6b-753">&lt;optional&gt;</span></span> | <span data-ttu-id="e3f6b-754">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-754">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="e3f6b-755">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="e3f6b-755">Object</span></span> | <span data-ttu-id="e3f6b-756">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="e3f6b-756">&lt;optional&gt;</span></span> | <span data-ttu-id="e3f6b-757">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-757">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="e3f6b-758">function</span><span class="sxs-lookup"><span data-stu-id="e3f6b-758">function</span></span>| <span data-ttu-id="e3f6b-759">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="e3f6b-759">&lt;optional&gt;</span></span>|<span data-ttu-id="e3f6b-760">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="e3f6b-760">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e3f6b-761">Requirements</span><span class="sxs-lookup"><span data-stu-id="e3f6b-761">Requirements</span></span>

|<span data-ttu-id="e3f6b-762">要件</span><span class="sxs-lookup"><span data-stu-id="e3f6b-762">Requirement</span></span>| <span data-ttu-id="e3f6b-763">値</span><span class="sxs-lookup"><span data-stu-id="e3f6b-763">Value</span></span>|
|---|---|
|[<span data-ttu-id="e3f6b-764">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e3f6b-764">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e3f6b-765">1.5</span><span class="sxs-lookup"><span data-stu-id="e3f6b-765">1.5</span></span> |
|[<span data-ttu-id="e3f6b-766">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e3f6b-766">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e3f6b-767">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e3f6b-767">ReadItem</span></span> |
|[<span data-ttu-id="e3f6b-768">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e3f6b-768">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e3f6b-769">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e3f6b-769">Compose or Read</span></span>|
