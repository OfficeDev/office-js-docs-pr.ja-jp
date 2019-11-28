---
title: Office. メールボックス要件セット1.7
description: ''
ms.date: 11/27/2019
localization_priority: Normal
ms.openlocfilehash: c310ad38bb9821955fb0571d3693ce39715376f4
ms.sourcegitcommit: 05a883a7fd89136301ce35aabc57638e9f563288
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/27/2019
ms.locfileid: "39629673"
---
# <a name="mailbox"></a><span data-ttu-id="e7022-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="e7022-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="e7022-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="e7022-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="e7022-104">Microsoft Outlook の Outlook アドイン オブジェクト モデルへのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="e7022-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="e7022-105">要件</span><span class="sxs-lookup"><span data-stu-id="e7022-105">Requirements</span></span>

|<span data-ttu-id="e7022-106">要件</span><span class="sxs-lookup"><span data-stu-id="e7022-106">Requirement</span></span>| <span data-ttu-id="e7022-107">値</span><span class="sxs-lookup"><span data-stu-id="e7022-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7022-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e7022-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e7022-109">1.0</span><span class="sxs-lookup"><span data-stu-id="e7022-109">1.0</span></span>|
|[<span data-ttu-id="e7022-110">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e7022-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e7022-111">制限あり</span><span class="sxs-lookup"><span data-stu-id="e7022-111">Restricted</span></span>|
|[<span data-ttu-id="e7022-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e7022-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e7022-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e7022-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="e7022-114">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="e7022-114">Members and methods</span></span>

| <span data-ttu-id="e7022-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="e7022-115">Member</span></span> | <span data-ttu-id="e7022-116">種類</span><span class="sxs-lookup"><span data-stu-id="e7022-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="e7022-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="e7022-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="e7022-118">メンバー</span><span class="sxs-lookup"><span data-stu-id="e7022-118">Member</span></span> |
| [<span data-ttu-id="e7022-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="e7022-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="e7022-120">メンバー</span><span class="sxs-lookup"><span data-stu-id="e7022-120">Member</span></span> |
| [<span data-ttu-id="e7022-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="e7022-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="e7022-122">メソッド</span><span class="sxs-lookup"><span data-stu-id="e7022-122">Method</span></span> |
| [<span data-ttu-id="e7022-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="e7022-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="e7022-124">メソッド</span><span class="sxs-lookup"><span data-stu-id="e7022-124">Method</span></span> |
| [<span data-ttu-id="e7022-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="e7022-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="e7022-126">メソッド</span><span class="sxs-lookup"><span data-stu-id="e7022-126">Method</span></span> |
| [<span data-ttu-id="e7022-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="e7022-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="e7022-128">メソッド</span><span class="sxs-lookup"><span data-stu-id="e7022-128">Method</span></span> |
| [<span data-ttu-id="e7022-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="e7022-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="e7022-130">メソッド</span><span class="sxs-lookup"><span data-stu-id="e7022-130">Method</span></span> |
| [<span data-ttu-id="e7022-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="e7022-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="e7022-132">メソッド</span><span class="sxs-lookup"><span data-stu-id="e7022-132">Method</span></span> |
| [<span data-ttu-id="e7022-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="e7022-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="e7022-134">メソッド</span><span class="sxs-lookup"><span data-stu-id="e7022-134">Method</span></span> |
| [<span data-ttu-id="e7022-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="e7022-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="e7022-136">メソッド</span><span class="sxs-lookup"><span data-stu-id="e7022-136">Method</span></span> |
| [<span data-ttu-id="e7022-137">displayNewMessageForm</span><span class="sxs-lookup"><span data-stu-id="e7022-137">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="e7022-138">メソッド</span><span class="sxs-lookup"><span data-stu-id="e7022-138">Method</span></span> |
| [<span data-ttu-id="e7022-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="e7022-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="e7022-140">メソッド</span><span class="sxs-lookup"><span data-stu-id="e7022-140">Method</span></span> |
| [<span data-ttu-id="e7022-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="e7022-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="e7022-142">メソッド</span><span class="sxs-lookup"><span data-stu-id="e7022-142">Method</span></span> |
| [<span data-ttu-id="e7022-143">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="e7022-143">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="e7022-144">メソッド</span><span class="sxs-lookup"><span data-stu-id="e7022-144">Method</span></span> |
| [<span data-ttu-id="e7022-145">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="e7022-145">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="e7022-146">メソッド</span><span class="sxs-lookup"><span data-stu-id="e7022-146">Method</span></span> |
| [<span data-ttu-id="e7022-147">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="e7022-147">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="e7022-148">メソッド</span><span class="sxs-lookup"><span data-stu-id="e7022-148">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="e7022-149">名前空間</span><span class="sxs-lookup"><span data-stu-id="e7022-149">Namespaces</span></span>

<span data-ttu-id="e7022-150">[diagnostics](Office.context.mailbox.diagnostics.md):Outlook アドインに診断情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="e7022-150">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="e7022-151">[item](Office.context.mailbox.item.md):Outlook アドインのメッセージや予定にアクセスするためのメソッドとプロパティを提供します。</span><span class="sxs-lookup"><span data-stu-id="e7022-151">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="e7022-152">[userProfile](Office.context.mailbox.userProfile.md):Outlook アドインのユーザーに関する情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="e7022-152">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="e7022-153">Members</span><span class="sxs-lookup"><span data-stu-id="e7022-153">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="e7022-154">ewsUrl: String</span><span class="sxs-lookup"><span data-stu-id="e7022-154">ewsUrl: String</span></span>

<span data-ttu-id="e7022-p101">このメール アカウントの Exchange Web サービス (EWS) エンドポイントの URL を取得します。読み取りモードのみです。</span><span class="sxs-lookup"><span data-stu-id="e7022-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="e7022-157">このメンバーは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="e7022-157">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="e7022-p102">
  `ewsUrl\` 値は、リモート サービスで、ユーザーのメールボックスに EWS 呼び出しを行うために使うことができます。たとえば、[選択したアイテムから添付ファイルを取得する](/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="e7022-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="e7022-160">アプリが閲覧モードで `ewsUrl` メンバーを呼び出すには、アプリのマニフェスト内に **ReadItem** アクセス許可が指定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="e7022-160">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="e7022-p103">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出してから、`ewsUrl` メンバーを使用する必要があります。アプリには、`saveAsync` メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="e7022-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="e7022-163">型</span><span class="sxs-lookup"><span data-stu-id="e7022-163">Type</span></span>

*   <span data-ttu-id="e7022-164">String</span><span class="sxs-lookup"><span data-stu-id="e7022-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e7022-165">Requirements</span><span class="sxs-lookup"><span data-stu-id="e7022-165">Requirements</span></span>

|<span data-ttu-id="e7022-166">要件</span><span class="sxs-lookup"><span data-stu-id="e7022-166">Requirement</span></span>| <span data-ttu-id="e7022-167">値</span><span class="sxs-lookup"><span data-stu-id="e7022-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7022-168">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e7022-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e7022-169">1.0</span><span class="sxs-lookup"><span data-stu-id="e7022-169">1.0</span></span>|
|[<span data-ttu-id="e7022-170">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e7022-170">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e7022-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e7022-171">ReadItem</span></span>|
|[<span data-ttu-id="e7022-172">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e7022-172">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e7022-173">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e7022-173">Compose or Read</span></span>|

<br>

---
---

#### <a name="resturl-string"></a><span data-ttu-id="e7022-174">restUrl: String</span><span class="sxs-lookup"><span data-stu-id="e7022-174">restUrl: String</span></span>

<span data-ttu-id="e7022-175">この電子メール アカウントの REST エンドポイントの URL を取得します。</span><span class="sxs-lookup"><span data-stu-id="e7022-175">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="e7022-176">`restUrl` 値は、ユーザーのメールボックスに [REST API](/outlook/rest/) 呼び出しを行うために使用できます。</span><span class="sxs-lookup"><span data-stu-id="e7022-176">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="e7022-177">型</span><span class="sxs-lookup"><span data-stu-id="e7022-177">Type</span></span>

*   <span data-ttu-id="e7022-178">String</span><span class="sxs-lookup"><span data-stu-id="e7022-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e7022-179">Requirements</span><span class="sxs-lookup"><span data-stu-id="e7022-179">Requirements</span></span>

|<span data-ttu-id="e7022-180">要件</span><span class="sxs-lookup"><span data-stu-id="e7022-180">Requirement</span></span>| <span data-ttu-id="e7022-181">値</span><span class="sxs-lookup"><span data-stu-id="e7022-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7022-182">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e7022-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e7022-183">1.5</span><span class="sxs-lookup"><span data-stu-id="e7022-183">1.5</span></span> |
|[<span data-ttu-id="e7022-184">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e7022-184">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e7022-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e7022-185">ReadItem</span></span>|
|[<span data-ttu-id="e7022-186">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e7022-186">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e7022-187">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e7022-187">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="e7022-188">メソッド</span><span class="sxs-lookup"><span data-stu-id="e7022-188">Methods</span></span>

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="e7022-189">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="e7022-189">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="e7022-190">サポートされているイベントのイベント ハンドラーを追加します。</span><span class="sxs-lookup"><span data-stu-id="e7022-190">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="e7022-191">現在、サポートされているイベントの種類は `Office.EventType.ItemChanged` だけです。</span><span class="sxs-lookup"><span data-stu-id="e7022-191">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e7022-192">パラメーター</span><span class="sxs-lookup"><span data-stu-id="e7022-192">Parameters</span></span>

| <span data-ttu-id="e7022-193">名前</span><span class="sxs-lookup"><span data-stu-id="e7022-193">Name</span></span> | <span data-ttu-id="e7022-194">種類</span><span class="sxs-lookup"><span data-stu-id="e7022-194">Type</span></span> | <span data-ttu-id="e7022-195">属性</span><span class="sxs-lookup"><span data-stu-id="e7022-195">Attributes</span></span> | <span data-ttu-id="e7022-196">説明</span><span class="sxs-lookup"><span data-stu-id="e7022-196">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="e7022-197">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="e7022-197">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="e7022-198">ハンドラーを呼び出す必要のあるイベント。</span><span class="sxs-lookup"><span data-stu-id="e7022-198">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="e7022-199">Function</span><span class="sxs-lookup"><span data-stu-id="e7022-199">Function</span></span> || <span data-ttu-id="e7022-p104">イベントを処理する関数。関数は、オブジェクト リテラルである単一パラメーターを受け入れる必要があります。パラメーターの `type` プロパティは、`addHandlerAsync` に渡される `eventType` パラメーターと一致します。</span><span class="sxs-lookup"><span data-stu-id="e7022-p104">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="e7022-203">Object</span><span class="sxs-lookup"><span data-stu-id="e7022-203">Object</span></span> | <span data-ttu-id="e7022-204">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="e7022-204">&lt;optional&gt;</span></span> | <span data-ttu-id="e7022-205">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="e7022-205">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="e7022-206">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="e7022-206">Object</span></span> | <span data-ttu-id="e7022-207">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="e7022-207">&lt;optional&gt;</span></span> | <span data-ttu-id="e7022-208">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="e7022-208">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="e7022-209">function</span><span class="sxs-lookup"><span data-stu-id="e7022-209">function</span></span>| <span data-ttu-id="e7022-210">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="e7022-210">&lt;optional&gt;</span></span>|<span data-ttu-id="e7022-211">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="e7022-211">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e7022-212">Requirements</span><span class="sxs-lookup"><span data-stu-id="e7022-212">Requirements</span></span>

|<span data-ttu-id="e7022-213">要件</span><span class="sxs-lookup"><span data-stu-id="e7022-213">Requirement</span></span>| <span data-ttu-id="e7022-214">値</span><span class="sxs-lookup"><span data-stu-id="e7022-214">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7022-215">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e7022-215">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e7022-216">1.5</span><span class="sxs-lookup"><span data-stu-id="e7022-216">1.5</span></span> |
|[<span data-ttu-id="e7022-217">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e7022-217">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e7022-218">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e7022-218">ReadItem</span></span> |
|[<span data-ttu-id="e7022-219">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e7022-219">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e7022-220">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e7022-220">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e7022-221">例</span><span class="sxs-lookup"><span data-stu-id="e7022-221">Example</span></span>

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
};
```

<br>

---
---

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="e7022-222">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="e7022-222">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="e7022-223">REST 形式のアイテム ID を EWS 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="e7022-223">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="e7022-224">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="e7022-224">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="e7022-p105">REST API ([Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) や [Microsoft Graph](https://graph.microsoft.io/) など) で取得されたアイテム ID は、Exchange Web サービス (EWS) に使用される形式とは異なる形式を使用します。`convertToEwsId` メソッドは、REST 形式の ID を EWS 用の適切な形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="e7022-p105">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e7022-227">パラメーター</span><span class="sxs-lookup"><span data-stu-id="e7022-227">Parameters</span></span>

|<span data-ttu-id="e7022-228">名前</span><span class="sxs-lookup"><span data-stu-id="e7022-228">Name</span></span>| <span data-ttu-id="e7022-229">種類</span><span class="sxs-lookup"><span data-stu-id="e7022-229">Type</span></span>| <span data-ttu-id="e7022-230">説明</span><span class="sxs-lookup"><span data-stu-id="e7022-230">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="e7022-231">String</span><span class="sxs-lookup"><span data-stu-id="e7022-231">String</span></span>|<span data-ttu-id="e7022-232">Outlook REST API 形式のアイテム ID</span><span class="sxs-lookup"><span data-stu-id="e7022-232">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="e7022-233">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="e7022-233">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.7)|<span data-ttu-id="e7022-234">アイテム ID の取得に使用された Outlook REST API のバージョンを示す値。</span><span class="sxs-lookup"><span data-stu-id="e7022-234">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e7022-235">要件</span><span class="sxs-lookup"><span data-stu-id="e7022-235">Requirements</span></span>

|<span data-ttu-id="e7022-236">要件</span><span class="sxs-lookup"><span data-stu-id="e7022-236">Requirement</span></span>| <span data-ttu-id="e7022-237">値</span><span class="sxs-lookup"><span data-stu-id="e7022-237">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7022-238">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e7022-238">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e7022-239">1.3</span><span class="sxs-lookup"><span data-stu-id="e7022-239">1.3</span></span>|
|[<span data-ttu-id="e7022-240">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e7022-240">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e7022-241">制限あり</span><span class="sxs-lookup"><span data-stu-id="e7022-241">Restricted</span></span>|
|[<span data-ttu-id="e7022-242">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e7022-242">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e7022-243">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e7022-243">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e7022-244">戻り値:</span><span class="sxs-lookup"><span data-stu-id="e7022-244">Returns:</span></span>

<span data-ttu-id="e7022-245">型:String</span><span class="sxs-lookup"><span data-stu-id="e7022-245">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="e7022-246">例</span><span class="sxs-lookup"><span data-stu-id="e7022-246">Example</span></span>

```js
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-17"></a><span data-ttu-id="e7022-247">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.7)}</span><span class="sxs-lookup"><span data-stu-id="e7022-247">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.7)}</span></span>

<span data-ttu-id="e7022-248">クライアントのローカル時間で時間情報が含まれている辞書を取得します。</span><span class="sxs-lookup"><span data-stu-id="e7022-248">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="e7022-p106">Outlook on the web または Outlook デスクトップのメールアプリでは、日付と時刻に異なるタイムゾーンを使用できます。Outlook デスクトップは、クライアント コンピューターのタイム ゾーンを使用します。Outlook on the web は、Exchange 管理センター (EAC) で設定されたタイム ゾーンを使用します。日付と時刻の値は、ユーザー インターフェイスに表示される値が、常にユーザーが期待するタイム ゾーンと一致するように処理する必要があります。</span><span class="sxs-lookup"><span data-stu-id="e7022-p106">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times. Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="e7022-p107">Outlook デスクトップ クライアントでメール アプリを実行している場合、`convertToLocalClientTime` メソッドは、クライアント コンピューターのタイム ゾーンに設定された値のディクショナリ オブジェクトを返します。Outlook on the web でメール アプリを実行している場合、`convertToLocalClientTime` メソッドは、EAC で指定したタイム ゾーンに設定された値のディクショナリ オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="e7022-p107">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e7022-254">パラメーター</span><span class="sxs-lookup"><span data-stu-id="e7022-254">Parameters</span></span>

|<span data-ttu-id="e7022-255">名前</span><span class="sxs-lookup"><span data-stu-id="e7022-255">Name</span></span>| <span data-ttu-id="e7022-256">種類</span><span class="sxs-lookup"><span data-stu-id="e7022-256">Type</span></span>| <span data-ttu-id="e7022-257">説明</span><span class="sxs-lookup"><span data-stu-id="e7022-257">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="e7022-258">日付</span><span class="sxs-lookup"><span data-stu-id="e7022-258">Date</span></span>|<span data-ttu-id="e7022-259">日付オブジェクト</span><span class="sxs-lookup"><span data-stu-id="e7022-259">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e7022-260">Requirements</span><span class="sxs-lookup"><span data-stu-id="e7022-260">Requirements</span></span>

|<span data-ttu-id="e7022-261">要件</span><span class="sxs-lookup"><span data-stu-id="e7022-261">Requirement</span></span>| <span data-ttu-id="e7022-262">値</span><span class="sxs-lookup"><span data-stu-id="e7022-262">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7022-263">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e7022-263">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e7022-264">1.0</span><span class="sxs-lookup"><span data-stu-id="e7022-264">1.0</span></span>|
|[<span data-ttu-id="e7022-265">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e7022-265">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e7022-266">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e7022-266">ReadItem</span></span>|
|[<span data-ttu-id="e7022-267">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e7022-267">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e7022-268">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e7022-268">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e7022-269">戻り値:</span><span class="sxs-lookup"><span data-stu-id="e7022-269">Returns:</span></span>

<span data-ttu-id="e7022-270">型:[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="e7022-270">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.7)</span></span>

<br>

---
---

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="e7022-271">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="e7022-271">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="e7022-272">EWS 形式のアイテム ID を REST 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="e7022-272">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="e7022-273">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="e7022-273">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="e7022-p108">EWS または `itemId` プロパティで取得されるアイテム ID は、REST API ([Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) や [Microsoft Graph](https://graph.microsoft.io/) など) に使用される形式とは異なる形式を使用します。`convertToRestId` メソッドは、EWS 形式の ID を REST 用の適切な形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="e7022-p108">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e7022-276">パラメーター</span><span class="sxs-lookup"><span data-stu-id="e7022-276">Parameters</span></span>

|<span data-ttu-id="e7022-277">名前</span><span class="sxs-lookup"><span data-stu-id="e7022-277">Name</span></span>| <span data-ttu-id="e7022-278">型</span><span class="sxs-lookup"><span data-stu-id="e7022-278">Type</span></span>| <span data-ttu-id="e7022-279">説明</span><span class="sxs-lookup"><span data-stu-id="e7022-279">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="e7022-280">String</span><span class="sxs-lookup"><span data-stu-id="e7022-280">String</span></span>|<span data-ttu-id="e7022-281">Exchange Web サービス (EWS) 形式のアイテム ID</span><span class="sxs-lookup"><span data-stu-id="e7022-281">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="e7022-282">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="e7022-282">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.7)|<span data-ttu-id="e7022-283">変換後の ID を使用する Outlook REST API のバージョンを示す値。</span><span class="sxs-lookup"><span data-stu-id="e7022-283">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e7022-284">Requirements</span><span class="sxs-lookup"><span data-stu-id="e7022-284">Requirements</span></span>

|<span data-ttu-id="e7022-285">要件</span><span class="sxs-lookup"><span data-stu-id="e7022-285">Requirement</span></span>| <span data-ttu-id="e7022-286">値</span><span class="sxs-lookup"><span data-stu-id="e7022-286">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7022-287">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e7022-287">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e7022-288">1.3</span><span class="sxs-lookup"><span data-stu-id="e7022-288">1.3</span></span>|
|[<span data-ttu-id="e7022-289">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e7022-289">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e7022-290">制限あり</span><span class="sxs-lookup"><span data-stu-id="e7022-290">Restricted</span></span>|
|[<span data-ttu-id="e7022-291">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e7022-291">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e7022-292">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e7022-292">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e7022-293">戻り値:</span><span class="sxs-lookup"><span data-stu-id="e7022-293">Returns:</span></span>

<span data-ttu-id="e7022-294">型:String</span><span class="sxs-lookup"><span data-stu-id="e7022-294">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="e7022-295">例</span><span class="sxs-lookup"><span data-stu-id="e7022-295">Example</span></span>

```js
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="e7022-296">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="e7022-296">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="e7022-297">時間情報が含まれているディクショナリから日付オブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="e7022-297">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="e7022-298">`convertToUtcClientTime` メソッドは、ローカルの日付と時刻を含むディクショナリを、ローカルの日付と時刻の正しい値を持つ日付オブジェクトに変換します。</span><span class="sxs-lookup"><span data-stu-id="e7022-298">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e7022-299">パラメーター</span><span class="sxs-lookup"><span data-stu-id="e7022-299">Parameters</span></span>

|<span data-ttu-id="e7022-300">名前</span><span class="sxs-lookup"><span data-stu-id="e7022-300">Name</span></span>| <span data-ttu-id="e7022-301">型</span><span class="sxs-lookup"><span data-stu-id="e7022-301">Type</span></span>| <span data-ttu-id="e7022-302">説明</span><span class="sxs-lookup"><span data-stu-id="e7022-302">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="e7022-303">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="e7022-303">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.7)|<span data-ttu-id="e7022-304">変換するローカル時刻の値。</span><span class="sxs-lookup"><span data-stu-id="e7022-304">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e7022-305">Requirements</span><span class="sxs-lookup"><span data-stu-id="e7022-305">Requirements</span></span>

|<span data-ttu-id="e7022-306">要件</span><span class="sxs-lookup"><span data-stu-id="e7022-306">Requirement</span></span>| <span data-ttu-id="e7022-307">値</span><span class="sxs-lookup"><span data-stu-id="e7022-307">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7022-308">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e7022-308">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e7022-309">1.0</span><span class="sxs-lookup"><span data-stu-id="e7022-309">1.0</span></span>|
|[<span data-ttu-id="e7022-310">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e7022-310">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e7022-311">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e7022-311">ReadItem</span></span>|
|[<span data-ttu-id="e7022-312">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e7022-312">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e7022-313">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e7022-313">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e7022-314">戻り値:</span><span class="sxs-lookup"><span data-stu-id="e7022-314">Returns:</span></span>

<span data-ttu-id="e7022-315">時間が UTC で表現された日付オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="e7022-315">A Date object with the time expressed in UTC.</span></span>

<span data-ttu-id="e7022-316">型: Date</span><span class="sxs-lookup"><span data-stu-id="e7022-316">Type: Date</span></span>

##### <a name="example"></a><span data-ttu-id="e7022-317">例</span><span class="sxs-lookup"><span data-stu-id="e7022-317">Example</span></span>

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

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="e7022-318">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="e7022-318">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="e7022-319">既存の予定を表示します。</span><span class="sxs-lookup"><span data-stu-id="e7022-319">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="e7022-320">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="e7022-320">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="e7022-321">`displayAppointmentForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存の予定を開きます。</span><span class="sxs-lookup"><span data-stu-id="e7022-321">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="e7022-p109">Outlook on Mac では、このメソッドを使用して定期的な系列に含まれない単発の予定や定期的な系列のマスター予定を表示できます。ただし、系列のインスタンスは表示できません。これは、Outlook on Mac では定期的な系列のインスタンスのプロパティ (アイテム ID を含む) にアクセスできないためです。</span><span class="sxs-lookup"><span data-stu-id="e7022-p109">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="e7022-324">Outlook on the web では、このメソッドはフォームの本文が 32KB 以下の文字数の場合にのみ指定のフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="e7022-324">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="e7022-325">指定のアイテム識別子が既存の予定を表していない場合は、クライアント コンピューターまたはデバイスで空のウィンドウが開き、エラー メッセージは返されません。</span><span class="sxs-lookup"><span data-stu-id="e7022-325">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e7022-326">パラメーター</span><span class="sxs-lookup"><span data-stu-id="e7022-326">Parameters</span></span>

|<span data-ttu-id="e7022-327">名前</span><span class="sxs-lookup"><span data-stu-id="e7022-327">Name</span></span>| <span data-ttu-id="e7022-328">型</span><span class="sxs-lookup"><span data-stu-id="e7022-328">Type</span></span>| <span data-ttu-id="e7022-329">説明</span><span class="sxs-lookup"><span data-stu-id="e7022-329">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="e7022-330">String</span><span class="sxs-lookup"><span data-stu-id="e7022-330">String</span></span>|<span data-ttu-id="e7022-331">既存の予定の Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="e7022-331">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e7022-332">Requirements</span><span class="sxs-lookup"><span data-stu-id="e7022-332">Requirements</span></span>

|<span data-ttu-id="e7022-333">要件</span><span class="sxs-lookup"><span data-stu-id="e7022-333">Requirement</span></span>| <span data-ttu-id="e7022-334">値</span><span class="sxs-lookup"><span data-stu-id="e7022-334">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7022-335">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e7022-335">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e7022-336">1.0</span><span class="sxs-lookup"><span data-stu-id="e7022-336">1.0</span></span>|
|[<span data-ttu-id="e7022-337">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e7022-337">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e7022-338">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e7022-338">ReadItem</span></span>|
|[<span data-ttu-id="e7022-339">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e7022-339">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e7022-340">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e7022-340">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e7022-341">例</span><span class="sxs-lookup"><span data-stu-id="e7022-341">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

<br>

---
---

#### <a name="displaymessageformitemid"></a><span data-ttu-id="e7022-342">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="e7022-342">displayMessageForm(itemId)</span></span>

<span data-ttu-id="e7022-343">既存のメッセージを表示します。</span><span class="sxs-lookup"><span data-stu-id="e7022-343">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="e7022-344">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="e7022-344">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="e7022-345">`displayMessageForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存のメッセージを開きます。</span><span class="sxs-lookup"><span data-stu-id="e7022-345">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="e7022-346">Outlook on the web では、このメソッドはフォームの本文が 32 KB 以下の文字数の場合にのみ指定のフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="e7022-346">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="e7022-347">指定のアイテム識別子が既存のメッセージを表していない場合は、クライアント コンピューターにはメッセージは表示されず、エラー メッセージも返されません。</span><span class="sxs-lookup"><span data-stu-id="e7022-347">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="e7022-p110">予定を表す `itemId` を含む `displayMessageForm` を使用しないでください。既存の予定を表示するには、`displayAppointmentForm` メソッドを使用します。新しい予定を作成するフォームを表示するには、`displayNewAppointmentForm` メソッドを使用します。</span><span class="sxs-lookup"><span data-stu-id="e7022-p110">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e7022-350">パラメーター</span><span class="sxs-lookup"><span data-stu-id="e7022-350">Parameters</span></span>

|<span data-ttu-id="e7022-351">名前</span><span class="sxs-lookup"><span data-stu-id="e7022-351">Name</span></span>| <span data-ttu-id="e7022-352">型</span><span class="sxs-lookup"><span data-stu-id="e7022-352">Type</span></span>| <span data-ttu-id="e7022-353">説明</span><span class="sxs-lookup"><span data-stu-id="e7022-353">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="e7022-354">String</span><span class="sxs-lookup"><span data-stu-id="e7022-354">String</span></span>|<span data-ttu-id="e7022-355">既存のメッセージの Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="e7022-355">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e7022-356">要件</span><span class="sxs-lookup"><span data-stu-id="e7022-356">Requirements</span></span>

|<span data-ttu-id="e7022-357">要件</span><span class="sxs-lookup"><span data-stu-id="e7022-357">Requirement</span></span>| <span data-ttu-id="e7022-358">値</span><span class="sxs-lookup"><span data-stu-id="e7022-358">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7022-359">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e7022-359">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e7022-360">1.0</span><span class="sxs-lookup"><span data-stu-id="e7022-360">1.0</span></span>|
|[<span data-ttu-id="e7022-361">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e7022-361">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e7022-362">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e7022-362">ReadItem</span></span>|
|[<span data-ttu-id="e7022-363">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e7022-363">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e7022-364">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e7022-364">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e7022-365">例</span><span class="sxs-lookup"><span data-stu-id="e7022-365">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

<br>

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="e7022-366">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="e7022-366">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="e7022-367">新しい予定を作成するためのフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="e7022-367">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="e7022-368">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="e7022-368">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="e7022-p111">`displayNewAppointmentForm` メソッドを使用すると、ユーザーが新しい予定または会議を作成できるフォームが開きます。パラメーターを指定すると、予定のフォーム フィールドにパラメーターの内容が自動的に設定されます。</span><span class="sxs-lookup"><span data-stu-id="e7022-p111">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="e7022-p112">Outlook on the web およびモバイル デバイスでは、このメソッドは常に出席者フィールドが含まれるフォームを表示します。入力引数として出席者を指定しないと、このメソッドは **[保存]** ボタンのあるフォームを表示します。出席者を指定した場合には、フォームにその出席者と **[送信]** ボタンが表示されます。</span><span class="sxs-lookup"><span data-stu-id="e7022-p112">In Outlook on the web and mobile devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="e7022-p113">Outlook リッチ クライアントと Outlook RT で、`requiredAttendees`、`optionalAttendees`、または `resources` パラメーターに出席者またはリソースを指定し、このメソッドを実行すると、**[送信]** ボタンがある会議フォームが表示されます。受信者を指定せずにこのメソッドを実行すると、**[保存して閉じる]** ボタンがある予定フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="e7022-p113">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="e7022-376">パラメーターのいずれかが指定のサイズ制限を超える場合、または不明なパラメーター名が指定されている場合は、例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="e7022-376">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e7022-377">パラメーター</span><span class="sxs-lookup"><span data-stu-id="e7022-377">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="e7022-378">すべてのパラメーターは省略可能です。</span><span class="sxs-lookup"><span data-stu-id="e7022-378">All parameters are optional.</span></span>

|<span data-ttu-id="e7022-379">名前</span><span class="sxs-lookup"><span data-stu-id="e7022-379">Name</span></span>| <span data-ttu-id="e7022-380">種類</span><span class="sxs-lookup"><span data-stu-id="e7022-380">Type</span></span>| <span data-ttu-id="e7022-381">説明</span><span class="sxs-lookup"><span data-stu-id="e7022-381">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="e7022-382">Object</span><span class="sxs-lookup"><span data-stu-id="e7022-382">Object</span></span> | <span data-ttu-id="e7022-383">新しい予定を記述するパラメーターのディクショナリ。</span><span class="sxs-lookup"><span data-stu-id="e7022-383">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="e7022-384">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span><span class="sxs-lookup"><span data-stu-id="e7022-384">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span></span> | <span data-ttu-id="e7022-p114">予定に必要な各出席者について、メール アドレスを含む文字列の配列、または `EmailAddressDetails` オブジェクトを含む配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="e7022-p114">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="e7022-387">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span><span class="sxs-lookup"><span data-stu-id="e7022-387">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span></span> | <span data-ttu-id="e7022-p115">予定の各任意出席者について、メール アドレスを含む文字列の配列、または `EmailAddressDetails` オブジェクトを含む配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="e7022-p115">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="e7022-390">日付</span><span class="sxs-lookup"><span data-stu-id="e7022-390">Date</span></span> | <span data-ttu-id="e7022-391">予定の開始日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="e7022-391">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="e7022-392">日付</span><span class="sxs-lookup"><span data-stu-id="e7022-392">Date</span></span> | <span data-ttu-id="e7022-393">予定の終了日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="e7022-393">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="e7022-394">String</span><span class="sxs-lookup"><span data-stu-id="e7022-394">String</span></span> | <span data-ttu-id="e7022-p116">予定の場所を含む文字列。文字列は最大 255 文字に制限されます。</span><span class="sxs-lookup"><span data-stu-id="e7022-p116">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="e7022-397">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="e7022-397">Array.&lt;String&gt;</span></span> | <span data-ttu-id="e7022-p117">予定に必要なリソースを含む文字列の配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="e7022-p117">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="e7022-400">String</span><span class="sxs-lookup"><span data-stu-id="e7022-400">String</span></span> | <span data-ttu-id="e7022-p118">予定の件名を含む文字列です。文字列は最大 255 文字に制限されます。</span><span class="sxs-lookup"><span data-stu-id="e7022-p118">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="e7022-403">String</span><span class="sxs-lookup"><span data-stu-id="e7022-403">String</span></span> | <span data-ttu-id="e7022-p119">予定の本文。本文の内容は、最大サイズが 32 KB に制限されます。</span><span class="sxs-lookup"><span data-stu-id="e7022-p119">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="e7022-406">要件</span><span class="sxs-lookup"><span data-stu-id="e7022-406">Requirements</span></span>

|<span data-ttu-id="e7022-407">要件</span><span class="sxs-lookup"><span data-stu-id="e7022-407">Requirement</span></span>| <span data-ttu-id="e7022-408">値</span><span class="sxs-lookup"><span data-stu-id="e7022-408">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7022-409">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e7022-409">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e7022-410">1.0</span><span class="sxs-lookup"><span data-stu-id="e7022-410">1.0</span></span>|
|[<span data-ttu-id="e7022-411">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e7022-411">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e7022-412">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e7022-412">ReadItem</span></span>|
|[<span data-ttu-id="e7022-413">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e7022-413">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e7022-414">読み取り</span><span class="sxs-lookup"><span data-stu-id="e7022-414">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e7022-415">例</span><span class="sxs-lookup"><span data-stu-id="e7022-415">Example</span></span>

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

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="e7022-416">displayNewMessageForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="e7022-416">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="e7022-417">新しいメッセージを作成するためのフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="e7022-417">Displays a form for creating a new message.</span></span>

<span data-ttu-id="e7022-p120">The `displayNewMessageForm` method opens a form that enables the user to create a new message. If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span><span class="sxs-lookup"><span data-stu-id="e7022-p120">The `displayNewMessageForm` method opens a form that enables the user to create a new message. If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="e7022-420">パラメーターのいずれかが指定のサイズ制限を超える場合、または不明なパラメーター名が指定されている場合は、例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="e7022-420">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e7022-421">パラメーター</span><span class="sxs-lookup"><span data-stu-id="e7022-421">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="e7022-422">すべてのパラメーターは省略可能です。</span><span class="sxs-lookup"><span data-stu-id="e7022-422">All parameters are optional.</span></span>

|<span data-ttu-id="e7022-423">名前</span><span class="sxs-lookup"><span data-stu-id="e7022-423">Name</span></span>| <span data-ttu-id="e7022-424">種類</span><span class="sxs-lookup"><span data-stu-id="e7022-424">Type</span></span>| <span data-ttu-id="e7022-425">説明</span><span class="sxs-lookup"><span data-stu-id="e7022-425">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="e7022-426">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="e7022-426">Object</span></span> | <span data-ttu-id="e7022-427">新しいメッセージを記述するパラメータの辞書。</span><span class="sxs-lookup"><span data-stu-id="e7022-427">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="e7022-428">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span><span class="sxs-lookup"><span data-stu-id="e7022-428">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span></span> | <span data-ttu-id="e7022-p121">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line. The array is limited to a maximum of 100 entries.</span><span class="sxs-lookup"><span data-stu-id="e7022-p121">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="e7022-431">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span><span class="sxs-lookup"><span data-stu-id="e7022-431">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span></span> | <span data-ttu-id="e7022-p122">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line. The array is limited to a maximum of 100 entries.</span><span class="sxs-lookup"><span data-stu-id="e7022-p122">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="e7022-434">配列。&lt;文字列&gt; | 配列。&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span><span class="sxs-lookup"><span data-stu-id="e7022-434">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span></span> | <span data-ttu-id="e7022-p123">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line. The array is limited to a maximum of 100 entries.</span><span class="sxs-lookup"><span data-stu-id="e7022-p123">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="e7022-437">String</span><span class="sxs-lookup"><span data-stu-id="e7022-437">String</span></span> | <span data-ttu-id="e7022-p124">A string containing the subject of the message. The string is limited to a maximum of 255 characters.</span><span class="sxs-lookup"><span data-stu-id="e7022-p124">A string containing the subject of the message. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="e7022-440">String</span><span class="sxs-lookup"><span data-stu-id="e7022-440">String</span></span> | <span data-ttu-id="e7022-p125">The HTML body of the message. The body content is limited to a maximum size of 32 KB.</span><span class="sxs-lookup"><span data-stu-id="e7022-p125">The HTML body of the message. The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="e7022-443">配列。&lt;オブジェクト&gt;</span><span class="sxs-lookup"><span data-stu-id="e7022-443">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="e7022-444">添付ファイルまたは添付アイテムである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="e7022-444">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="e7022-445">String</span><span class="sxs-lookup"><span data-stu-id="e7022-445">String</span></span> | <span data-ttu-id="e7022-p126">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="e7022-p126">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="e7022-448">String</span><span class="sxs-lookup"><span data-stu-id="e7022-448">String</span></span> | <span data-ttu-id="e7022-449">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="e7022-449">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="e7022-450">文字列</span><span class="sxs-lookup"><span data-stu-id="e7022-450">String</span></span> | <span data-ttu-id="e7022-p127">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="e7022-p127">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="e7022-453">ブール値</span><span class="sxs-lookup"><span data-stu-id="e7022-453">Boolean</span></span> | <span data-ttu-id="e7022-p128">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="e7022-p128">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="e7022-456">String</span><span class="sxs-lookup"><span data-stu-id="e7022-456">String</span></span> | <span data-ttu-id="e7022-p129">Only used if `type` is set to `item`. The EWS item id of the existing e-mail you want to attach to the new message. This is a string up to 100 characters.</span><span class="sxs-lookup"><span data-stu-id="e7022-p129">Only used if `type` is set to `item`. The EWS item id of the existing e-mail you want to attach to the new message. This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="e7022-460">Requirements</span><span class="sxs-lookup"><span data-stu-id="e7022-460">Requirements</span></span>

|<span data-ttu-id="e7022-461">要件</span><span class="sxs-lookup"><span data-stu-id="e7022-461">Requirement</span></span>| <span data-ttu-id="e7022-462">値</span><span class="sxs-lookup"><span data-stu-id="e7022-462">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7022-463">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e7022-463">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e7022-464">1.6</span><span class="sxs-lookup"><span data-stu-id="e7022-464">1.6</span></span> |
|[<span data-ttu-id="e7022-465">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e7022-465">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e7022-466">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e7022-466">ReadItem</span></span>|
|[<span data-ttu-id="e7022-467">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e7022-467">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e7022-468">読み取り</span><span class="sxs-lookup"><span data-stu-id="e7022-468">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e7022-469">例</span><span class="sxs-lookup"><span data-stu-id="e7022-469">Example</span></span>

```js
Office.context.mailbox.displayNewMessageForm(
  {
    // Copy the To line from current item.
    toRecipients: Office.context.mailbox.item.to
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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="e7022-470">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="e7022-470">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="e7022-471">REST API または Exchange Web サービスを呼び出すために使用するトークンを含む文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="e7022-471">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="e7022-p130">`getCallbackTokenAsync` メソッドは、ユーザーのメールボックスをホストする Exchange Server から不透明なトークンを取得する非同期の呼び出しを行います。コールバック トークンの有効期間は 5 分です。</span><span class="sxs-lookup"><span data-stu-id="e7022-p130">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="e7022-474">可能であれば、アドインでは Exchange Web サービスの代わりに REST API を使用することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="e7022-474">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span>

<span data-ttu-id="e7022-475">閲覧モードで `getCallbackTokenAsync` メソッドを呼び出すには、**ReadItem** の最小限のアクセス許可レベルが必要です。</span><span class="sxs-lookup"><span data-stu-id="e7022-475">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="e7022-476">作成モードで `getCallbackTokenAsync` を呼び出すには、アイテムを保存しておく必要があります。</span><span class="sxs-lookup"><span data-stu-id="e7022-476">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="e7022-477">[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドは、**ReadWriteItem** の最小限のアクセス許可レベルが必要です。</span><span class="sxs-lookup"><span data-stu-id="e7022-477">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

<span data-ttu-id="e7022-478">**REST トークン**</span><span class="sxs-lookup"><span data-stu-id="e7022-478">**REST Tokens**</span></span>

<span data-ttu-id="e7022-p132">REST トークンが要求された場合 (`options.isRest = true`)、結果トークンは Exchange Web サービスの呼び出しを認証するためには機能しません。アドインがマニフェストで [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) アクセス許可を指定していない限り、トークンの範囲は現在のアイテムとその添付ファイルへの読み取り専用アクセスに制限されます。`ReadWriteMailbox` アクセス許可が指定されている場合は、結果トークンは、メールを送信する機能など、メール、カレンダー、連絡先への読み取り/書き込みアクセスを付与します。</span><span class="sxs-lookup"><span data-stu-id="e7022-p132">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="e7022-482">アドインでは、`restUrl` プロパティを使用して、REST API 呼び出しを行うときに使用する正しい URL を決定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="e7022-482">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="e7022-483">**EWS トークン**</span><span class="sxs-lookup"><span data-stu-id="e7022-483">**EWS Tokens**</span></span>

<span data-ttu-id="e7022-p133">EWS トークンが要求された場合 (`options.isRest = false`)、結果トークンは REST API 呼び出しを認証するためには機能しません。トークンの範囲は、現在のアイテムへのアクセスに制限されます。</span><span class="sxs-lookup"><span data-stu-id="e7022-p133">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="e7022-486">アドインでは、`ewsUrl` プロパティを使用して、EWS 呼び出しを行うときに使用する正しい URL を決定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="e7022-486">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

<span data-ttu-id="e7022-487">トークンと、添付ファイル識別子またはアイテム識別子の両方をサードパーティ システムに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="e7022-487">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="e7022-488">サードパーティ システムは、トークンをベアラー承認トークンとして使用し、Exchange Web サービス (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) 操作または [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) 操作を呼び出して、添付ファイルまたはアイテムを返します。</span><span class="sxs-lookup"><span data-stu-id="e7022-488">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item.</span></span> <span data-ttu-id="e7022-489">たとえば、[選択したアイテムから添付ファイルを取得する](/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="e7022-489">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

##### <a name="parameters"></a><span data-ttu-id="e7022-490">パラメーター</span><span class="sxs-lookup"><span data-stu-id="e7022-490">Parameters</span></span>

|<span data-ttu-id="e7022-491">名前</span><span class="sxs-lookup"><span data-stu-id="e7022-491">Name</span></span>| <span data-ttu-id="e7022-492">型</span><span class="sxs-lookup"><span data-stu-id="e7022-492">Type</span></span>| <span data-ttu-id="e7022-493">属性</span><span class="sxs-lookup"><span data-stu-id="e7022-493">Attributes</span></span>| <span data-ttu-id="e7022-494">説明</span><span class="sxs-lookup"><span data-stu-id="e7022-494">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="e7022-495">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="e7022-495">Object</span></span> | <span data-ttu-id="e7022-496">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="e7022-496">&lt;optional&gt;</span></span> | <span data-ttu-id="e7022-497">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="e7022-497">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="e7022-498">ブール値</span><span class="sxs-lookup"><span data-stu-id="e7022-498">Boolean</span></span> |  <span data-ttu-id="e7022-499">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="e7022-499">&lt;optional&gt;</span></span> | <span data-ttu-id="e7022-p135">提供されたトークンを Outlook REST API または Exchange Web サービスに使用するかどうかを決定します。既定値は、`false` です。</span><span class="sxs-lookup"><span data-stu-id="e7022-p135">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="e7022-502">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="e7022-502">Object</span></span> |  <span data-ttu-id="e7022-503">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="e7022-503">&lt;optional&gt;</span></span> | <span data-ttu-id="e7022-504">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="e7022-504">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="e7022-505">function</span><span class="sxs-lookup"><span data-stu-id="e7022-505">function</span></span>||<span data-ttu-id="e7022-506">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="e7022-506">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="e7022-507">トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="e7022-507">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="e7022-508">エラーが発生した場合、 `asyncResult.error` および `asyncResult.diagnostics` のプロパティで追加情報が提供される場合があります。</span><span class="sxs-lookup"><span data-stu-id="e7022-508">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|

##### <a name="errors"></a><span data-ttu-id="e7022-509">エラー</span><span class="sxs-lookup"><span data-stu-id="e7022-509">Errors</span></span>

|<span data-ttu-id="e7022-510">エラー コード</span><span class="sxs-lookup"><span data-stu-id="e7022-510">Error code</span></span>|<span data-ttu-id="e7022-511">説明</span><span class="sxs-lookup"><span data-stu-id="e7022-511">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="e7022-512">要求が失敗しました。</span><span class="sxs-lookup"><span data-stu-id="e7022-512">The request has failed.</span></span> <span data-ttu-id="e7022-513">HTTP エラーコードの diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="e7022-513">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="e7022-514">Exchange サーバーがエラーを返しました。</span><span class="sxs-lookup"><span data-stu-id="e7022-514">The Exchange server returned an error.</span></span> <span data-ttu-id="e7022-515">詳細については、diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="e7022-515">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="e7022-516">ユーザーはネットワークに接続されていません。</span><span class="sxs-lookup"><span data-stu-id="e7022-516">The user is no longer connected to the network.</span></span> <span data-ttu-id="e7022-517">ネットワーク接続を確認し、やり直してください。</span><span class="sxs-lookup"><span data-stu-id="e7022-517">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e7022-518">要件</span><span class="sxs-lookup"><span data-stu-id="e7022-518">Requirements</span></span>

|<span data-ttu-id="e7022-519">必要条件</span><span class="sxs-lookup"><span data-stu-id="e7022-519">Requirement</span></span>| <span data-ttu-id="e7022-520">値</span><span class="sxs-lookup"><span data-stu-id="e7022-520">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7022-521">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e7022-521">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e7022-522">1.5</span><span class="sxs-lookup"><span data-stu-id="e7022-522">1.5</span></span> |
|[<span data-ttu-id="e7022-523">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e7022-523">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e7022-524">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e7022-524">ReadItem</span></span>|
|[<span data-ttu-id="e7022-525">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e7022-525">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e7022-526">新規作成と閲覧</span><span class="sxs-lookup"><span data-stu-id="e7022-526">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="e7022-527">例</span><span class="sxs-lookup"><span data-stu-id="e7022-527">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="e7022-528">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="e7022-528">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="e7022-529">Exchange Server から添付ファイルやアイテムを取得するために使うトークンを含む文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="e7022-529">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="e7022-p139">`getCallbackTokenAsync` メソッドは、ユーザーのメールボックスをホストする Exchange Server から不透明なトークンを取得する非同期の呼び出しを行います。コールバック トークンの有効期間は 5 分です。</span><span class="sxs-lookup"><span data-stu-id="e7022-p139">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="e7022-532">トークンと、添付ファイル識別子またはアイテム識別子の両方をサードパーティ システムに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="e7022-532">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="e7022-533">サードパーティ システムは、トークンをベアラー承認トークンとして使用し、Exchange Web サービス (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) 操作または [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) 操作を呼び出して、添付ファイルまたはアイテムを返します。</span><span class="sxs-lookup"><span data-stu-id="e7022-533">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item.</span></span> <span data-ttu-id="e7022-534">たとえば、[選択したアイテムから添付ファイルを取得する](/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="e7022-534">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="e7022-535">閲覧モードで `getCallbackTokenAsync` メソッドを呼び出すには、**ReadItem** の最小限のアクセス許可レベルが必要です。</span><span class="sxs-lookup"><span data-stu-id="e7022-535">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="e7022-536">作成モードで `getCallbackTokenAsync` を呼び出すには、アイテムを保存しておく必要があります。</span><span class="sxs-lookup"><span data-stu-id="e7022-536">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="e7022-537">[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドは、**ReadWriteItem** の最小限のアクセス許可レベルが必要です。</span><span class="sxs-lookup"><span data-stu-id="e7022-537">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e7022-538">パラメーター</span><span class="sxs-lookup"><span data-stu-id="e7022-538">Parameters</span></span>

|<span data-ttu-id="e7022-539">名前</span><span class="sxs-lookup"><span data-stu-id="e7022-539">Name</span></span>| <span data-ttu-id="e7022-540">型</span><span class="sxs-lookup"><span data-stu-id="e7022-540">Type</span></span>| <span data-ttu-id="e7022-541">属性</span><span class="sxs-lookup"><span data-stu-id="e7022-541">Attributes</span></span>| <span data-ttu-id="e7022-542">説明</span><span class="sxs-lookup"><span data-stu-id="e7022-542">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="e7022-543">function</span><span class="sxs-lookup"><span data-stu-id="e7022-543">function</span></span>||<span data-ttu-id="e7022-544">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="e7022-544">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="e7022-545">トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="e7022-545">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="e7022-546">エラーが発生した場合、 `asyncResult.error` および `asyncResult.diagnostics` のプロパティで追加情報が提供される場合があります。</span><span class="sxs-lookup"><span data-stu-id="e7022-546">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="e7022-547">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="e7022-547">Object</span></span>| <span data-ttu-id="e7022-548">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="e7022-548">&lt;optional&gt;</span></span>|<span data-ttu-id="e7022-549">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="e7022-549">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="e7022-550">エラー</span><span class="sxs-lookup"><span data-stu-id="e7022-550">Errors</span></span>

|<span data-ttu-id="e7022-551">エラー コード</span><span class="sxs-lookup"><span data-stu-id="e7022-551">Error code</span></span>|<span data-ttu-id="e7022-552">説明</span><span class="sxs-lookup"><span data-stu-id="e7022-552">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="e7022-553">要求が失敗しました。</span><span class="sxs-lookup"><span data-stu-id="e7022-553">The request has failed.</span></span> <span data-ttu-id="e7022-554">HTTP エラーコードの diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="e7022-554">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="e7022-555">Exchange サーバーがエラーを返しました。</span><span class="sxs-lookup"><span data-stu-id="e7022-555">The Exchange server returned an error.</span></span> <span data-ttu-id="e7022-556">詳細については、diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="e7022-556">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="e7022-557">ユーザーはネットワークに接続されていません。</span><span class="sxs-lookup"><span data-stu-id="e7022-557">The user is no longer connected to the network.</span></span> <span data-ttu-id="e7022-558">ネットワーク接続を確認し、やり直してください。</span><span class="sxs-lookup"><span data-stu-id="e7022-558">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e7022-559">要件</span><span class="sxs-lookup"><span data-stu-id="e7022-559">Requirements</span></span>

|<span data-ttu-id="e7022-560">要件</span><span class="sxs-lookup"><span data-stu-id="e7022-560">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="e7022-561">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e7022-561">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e7022-562">1.0</span><span class="sxs-lookup"><span data-stu-id="e7022-562">1.0</span></span> | <span data-ttu-id="e7022-563">1.3</span><span class="sxs-lookup"><span data-stu-id="e7022-563">1.3</span></span> |
|[<span data-ttu-id="e7022-564">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e7022-564">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e7022-565">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e7022-565">ReadItem</span></span> | <span data-ttu-id="e7022-566">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e7022-566">ReadItem</span></span> |
|[<span data-ttu-id="e7022-567">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e7022-567">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e7022-568">Read</span><span class="sxs-lookup"><span data-stu-id="e7022-568">Read</span></span> | <span data-ttu-id="e7022-569">Compose</span><span class="sxs-lookup"><span data-stu-id="e7022-569">Compose</span></span> |

##### <a name="example"></a><span data-ttu-id="e7022-570">例</span><span class="sxs-lookup"><span data-stu-id="e7022-570">Example</span></span>

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

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="e7022-571">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="e7022-571">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="e7022-572">ユーザーと Office アドインを識別するトークンを取得します。</span><span class="sxs-lookup"><span data-stu-id="e7022-572">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="e7022-573">`getUserIdentityTokenAsync` メソッドは、[アドインとユーザーをサード パーティのシステムで識別して認証](/outlook/add-ins/authentication)することのできるトークンを返します。</span><span class="sxs-lookup"><span data-stu-id="e7022-573">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="e7022-574">パラメーター</span><span class="sxs-lookup"><span data-stu-id="e7022-574">Parameters</span></span>

|<span data-ttu-id="e7022-575">名前</span><span class="sxs-lookup"><span data-stu-id="e7022-575">Name</span></span>| <span data-ttu-id="e7022-576">型</span><span class="sxs-lookup"><span data-stu-id="e7022-576">Type</span></span>| <span data-ttu-id="e7022-577">属性</span><span class="sxs-lookup"><span data-stu-id="e7022-577">Attributes</span></span>| <span data-ttu-id="e7022-578">説明</span><span class="sxs-lookup"><span data-stu-id="e7022-578">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="e7022-579">function</span><span class="sxs-lookup"><span data-stu-id="e7022-579">function</span></span>||<span data-ttu-id="e7022-580">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="e7022-580">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="e7022-581">トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="e7022-581">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="e7022-582">エラーが発生した場合、 `asyncResult.error` および `asyncResult.diagnostics` のプロパティで追加情報が提供される場合があります。</span><span class="sxs-lookup"><span data-stu-id="e7022-582">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="e7022-583">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="e7022-583">Object</span></span>| <span data-ttu-id="e7022-584">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="e7022-584">&lt;optional&gt;</span></span>|<span data-ttu-id="e7022-585">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="e7022-585">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="e7022-586">エラー</span><span class="sxs-lookup"><span data-stu-id="e7022-586">Errors</span></span>

|<span data-ttu-id="e7022-587">エラー コード</span><span class="sxs-lookup"><span data-stu-id="e7022-587">Error code</span></span>|<span data-ttu-id="e7022-588">説明</span><span class="sxs-lookup"><span data-stu-id="e7022-588">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="e7022-589">要求が失敗しました。</span><span class="sxs-lookup"><span data-stu-id="e7022-589">The request has failed.</span></span> <span data-ttu-id="e7022-590">HTTP エラーコードの diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="e7022-590">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="e7022-591">Exchange サーバーがエラーを返しました。</span><span class="sxs-lookup"><span data-stu-id="e7022-591">The Exchange server returned an error.</span></span> <span data-ttu-id="e7022-592">詳細については、diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="e7022-592">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="e7022-593">ユーザーはネットワークに接続されていません。</span><span class="sxs-lookup"><span data-stu-id="e7022-593">The user is no longer connected to the network.</span></span> <span data-ttu-id="e7022-594">ネットワーク接続を確認し、やり直してください。</span><span class="sxs-lookup"><span data-stu-id="e7022-594">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e7022-595">要件</span><span class="sxs-lookup"><span data-stu-id="e7022-595">Requirements</span></span>

|<span data-ttu-id="e7022-596">必要条件</span><span class="sxs-lookup"><span data-stu-id="e7022-596">Requirement</span></span>| <span data-ttu-id="e7022-597">値</span><span class="sxs-lookup"><span data-stu-id="e7022-597">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7022-598">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e7022-598">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e7022-599">1.0以降</span><span class="sxs-lookup"><span data-stu-id="e7022-599">1.0</span></span>|
|[<span data-ttu-id="e7022-600">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e7022-600">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e7022-601">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e7022-601">ReadItem</span></span>|
|[<span data-ttu-id="e7022-602">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e7022-602">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e7022-603">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e7022-603">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e7022-604">例</span><span class="sxs-lookup"><span data-stu-id="e7022-604">Example</span></span>

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

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="e7022-605">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="e7022-605">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="e7022-606">ユーザーのメールボックスをホストしている Exchange サーバー上の Exchange Web サービス (EWS) のサービスに対して非同期の要求を行います。</span><span class="sxs-lookup"><span data-stu-id="e7022-606">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="e7022-607">このメソッドは、次のシナリオではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="e7022-607">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="e7022-608">Outlook on iOS または Android の場合</span><span class="sxs-lookup"><span data-stu-id="e7022-608">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="e7022-609">アドインが Gmail のメールボックスに読み込まれる場合</span><span class="sxs-lookup"><span data-stu-id="e7022-609">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="e7022-610">このような場合は、アドインでは [REST API を使用](/outlook/add-ins/use-rest-api)して、代わりにユーザーのメールボックスにアクセスする必要があります。</span><span class="sxs-lookup"><span data-stu-id="e7022-610">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="e7022-p148">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange. See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span><span class="sxs-lookup"><span data-stu-id="e7022-p148">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange. See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="e7022-613">`makeEwsRequestAsync` メソッドでは、フォルダー関連アイテムを要求できません。</span><span class="sxs-lookup"><span data-stu-id="e7022-613">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="e7022-614">XML 要求では UTF-8 エンコードを指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="e7022-614">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="e7022-p149">`makeEwsRequestAsync` メソッドを使用するには、アドインに **ReadWriteMailbox** アクセス許可が必要です。**ReadWriteMailbox** アクセス許可と、`makeEwsRequestAsync` メソッドで呼び出せる EWS 操作の使い方については、「[ユーザーのメールボックスへのメール アドイン アクセスのアクセス許可を指定する](/outlook/add-ins/understanding-outlook-add-in-permissions)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e7022-p149">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="e7022-617">サーバー管理者は、クライアント アクセス サーバーの EWS ディレクトリで `OAuthAuthentication` を true に設定して、`makeEwsRequestAsync` メソッドで EWS 要求を行うことができるようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="e7022-617">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="e7022-618">バージョンの相違点</span><span class="sxs-lookup"><span data-stu-id="e7022-618">Version differences</span></span>

<span data-ttu-id="e7022-619">バージョン 15.0.4535.1004 より前のバージョンの Outlook で実行しているメール アプリで `makeEwsRequestAsync` メソッドを使う場合は、エンコード値を `ISO-8859-1` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="e7022-619">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="e7022-p150">Outlook on the web でメール アプリを実行している場合は、エンコード値を設定する必要はありません。mailbox.diagnostics.hostName プロパティを使って、メール アプリを Outlook で実行しているのか、Outlook on the web で実行しているのかを確認できます。mailbox.diagnostics.hostVersion プロパティを使って、どのバージョンの Outlook を使って実行しているかを確認できます。</span><span class="sxs-lookup"><span data-stu-id="e7022-p150">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e7022-623">パラメーター</span><span class="sxs-lookup"><span data-stu-id="e7022-623">Parameters</span></span>

|<span data-ttu-id="e7022-624">名前</span><span class="sxs-lookup"><span data-stu-id="e7022-624">Name</span></span>| <span data-ttu-id="e7022-625">型</span><span class="sxs-lookup"><span data-stu-id="e7022-625">Type</span></span>| <span data-ttu-id="e7022-626">属性</span><span class="sxs-lookup"><span data-stu-id="e7022-626">Attributes</span></span>| <span data-ttu-id="e7022-627">説明</span><span class="sxs-lookup"><span data-stu-id="e7022-627">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="e7022-628">String</span><span class="sxs-lookup"><span data-stu-id="e7022-628">String</span></span>||<span data-ttu-id="e7022-629">EWS 要求です。</span><span class="sxs-lookup"><span data-stu-id="e7022-629">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="e7022-630">function</span><span class="sxs-lookup"><span data-stu-id="e7022-630">function</span></span>||<span data-ttu-id="e7022-631">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="e7022-631">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="e7022-p151">The XML result of the EWS call is provided as a string in the `asyncResult.value` property. If the result exceeds 1 MB in size, an error message is returned instead.</span><span class="sxs-lookup"><span data-stu-id="e7022-p151">The XML result of the EWS call is provided as a string in the `asyncResult.value` property. If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="e7022-634">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="e7022-634">Object</span></span>| <span data-ttu-id="e7022-635">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="e7022-635">&lt;optional&gt;</span></span>|<span data-ttu-id="e7022-636">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="e7022-636">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e7022-637">要件</span><span class="sxs-lookup"><span data-stu-id="e7022-637">Requirements</span></span>

|<span data-ttu-id="e7022-638">要件</span><span class="sxs-lookup"><span data-stu-id="e7022-638">Requirement</span></span>| <span data-ttu-id="e7022-639">値</span><span class="sxs-lookup"><span data-stu-id="e7022-639">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7022-640">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e7022-640">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e7022-641">1.0</span><span class="sxs-lookup"><span data-stu-id="e7022-641">1.0</span></span>|
|[<span data-ttu-id="e7022-642">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e7022-642">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e7022-643">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="e7022-643">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="e7022-644">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e7022-644">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e7022-645">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e7022-645">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e7022-646">例</span><span class="sxs-lookup"><span data-stu-id="e7022-646">Example</span></span>

<span data-ttu-id="e7022-647">次の例は、`GetItem` 操作を使ってアイテムの件名を取得するため、`makeEwsRequestAsync` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="e7022-647">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="e7022-648">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="e7022-648">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="e7022-649">サポートされているイベントの種類のイベント ハンドラーを削除します。</span><span class="sxs-lookup"><span data-stu-id="e7022-649">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="e7022-650">現在、サポートされているイベントの種類は `Office.EventType.ItemChanged` だけです。</span><span class="sxs-lookup"><span data-stu-id="e7022-650">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e7022-651">パラメーター</span><span class="sxs-lookup"><span data-stu-id="e7022-651">Parameters</span></span>

| <span data-ttu-id="e7022-652">名前</span><span class="sxs-lookup"><span data-stu-id="e7022-652">Name</span></span> | <span data-ttu-id="e7022-653">型</span><span class="sxs-lookup"><span data-stu-id="e7022-653">Type</span></span> | <span data-ttu-id="e7022-654">属性</span><span class="sxs-lookup"><span data-stu-id="e7022-654">Attributes</span></span> | <span data-ttu-id="e7022-655">説明</span><span class="sxs-lookup"><span data-stu-id="e7022-655">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="e7022-656">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="e7022-656">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="e7022-657">ハンドラーを取り消すイベント。</span><span class="sxs-lookup"><span data-stu-id="e7022-657">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="e7022-658">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="e7022-658">Object</span></span> | <span data-ttu-id="e7022-659">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="e7022-659">&lt;optional&gt;</span></span> | <span data-ttu-id="e7022-660">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="e7022-660">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="e7022-661">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="e7022-661">Object</span></span> | <span data-ttu-id="e7022-662">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="e7022-662">&lt;optional&gt;</span></span> | <span data-ttu-id="e7022-663">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="e7022-663">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="e7022-664">function</span><span class="sxs-lookup"><span data-stu-id="e7022-664">function</span></span>| <span data-ttu-id="e7022-665">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="e7022-665">&lt;optional&gt;</span></span>|<span data-ttu-id="e7022-666">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="e7022-666">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e7022-667">Requirements</span><span class="sxs-lookup"><span data-stu-id="e7022-667">Requirements</span></span>

|<span data-ttu-id="e7022-668">要件</span><span class="sxs-lookup"><span data-stu-id="e7022-668">Requirement</span></span>| <span data-ttu-id="e7022-669">値</span><span class="sxs-lookup"><span data-stu-id="e7022-669">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7022-670">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e7022-670">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e7022-671">1.5</span><span class="sxs-lookup"><span data-stu-id="e7022-671">1.5</span></span> |
|[<span data-ttu-id="e7022-672">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e7022-672">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e7022-673">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e7022-673">ReadItem</span></span> |
|[<span data-ttu-id="e7022-674">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e7022-674">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e7022-675">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e7022-675">Compose or Read</span></span>|
