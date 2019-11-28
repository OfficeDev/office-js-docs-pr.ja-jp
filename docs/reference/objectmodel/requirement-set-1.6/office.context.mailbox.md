---
title: Office. メールボックス要件セット1.6
description: ''
ms.date: 11/27/2019
localization_priority: Normal
ms.openlocfilehash: 09c3930daf6f26edbc38b01f515ee5b1830ce802
ms.sourcegitcommit: 05a883a7fd89136301ce35aabc57638e9f563288
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/27/2019
ms.locfileid: "39629694"
---
# <a name="mailbox"></a><span data-ttu-id="43bc9-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="43bc9-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="43bc9-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="43bc9-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="43bc9-104">Microsoft Outlook の Outlook アドイン オブジェクト モデルへのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="43bc9-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="43bc9-105">要件</span><span class="sxs-lookup"><span data-stu-id="43bc9-105">Requirements</span></span>

|<span data-ttu-id="43bc9-106">要件</span><span class="sxs-lookup"><span data-stu-id="43bc9-106">Requirement</span></span>| <span data-ttu-id="43bc9-107">値</span><span class="sxs-lookup"><span data-stu-id="43bc9-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="43bc9-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="43bc9-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43bc9-109">1.0</span><span class="sxs-lookup"><span data-stu-id="43bc9-109">1.0</span></span>|
|[<span data-ttu-id="43bc9-110">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="43bc9-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43bc9-111">制限あり</span><span class="sxs-lookup"><span data-stu-id="43bc9-111">Restricted</span></span>|
|[<span data-ttu-id="43bc9-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="43bc9-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="43bc9-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="43bc9-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="43bc9-114">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="43bc9-114">Members and methods</span></span>

| <span data-ttu-id="43bc9-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="43bc9-115">Member</span></span> | <span data-ttu-id="43bc9-116">種類</span><span class="sxs-lookup"><span data-stu-id="43bc9-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="43bc9-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="43bc9-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="43bc9-118">メンバー</span><span class="sxs-lookup"><span data-stu-id="43bc9-118">Member</span></span> |
| [<span data-ttu-id="43bc9-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="43bc9-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="43bc9-120">メンバー</span><span class="sxs-lookup"><span data-stu-id="43bc9-120">Member</span></span> |
| [<span data-ttu-id="43bc9-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="43bc9-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="43bc9-122">メソッド</span><span class="sxs-lookup"><span data-stu-id="43bc9-122">Method</span></span> |
| [<span data-ttu-id="43bc9-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="43bc9-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="43bc9-124">メソッド</span><span class="sxs-lookup"><span data-stu-id="43bc9-124">Method</span></span> |
| [<span data-ttu-id="43bc9-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="43bc9-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="43bc9-126">メソッド</span><span class="sxs-lookup"><span data-stu-id="43bc9-126">Method</span></span> |
| [<span data-ttu-id="43bc9-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="43bc9-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="43bc9-128">メソッド</span><span class="sxs-lookup"><span data-stu-id="43bc9-128">Method</span></span> |
| [<span data-ttu-id="43bc9-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="43bc9-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="43bc9-130">メソッド</span><span class="sxs-lookup"><span data-stu-id="43bc9-130">Method</span></span> |
| [<span data-ttu-id="43bc9-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="43bc9-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="43bc9-132">メソッド</span><span class="sxs-lookup"><span data-stu-id="43bc9-132">Method</span></span> |
| [<span data-ttu-id="43bc9-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="43bc9-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="43bc9-134">メソッド</span><span class="sxs-lookup"><span data-stu-id="43bc9-134">Method</span></span> |
| [<span data-ttu-id="43bc9-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="43bc9-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="43bc9-136">メソッド</span><span class="sxs-lookup"><span data-stu-id="43bc9-136">Method</span></span> |
| [<span data-ttu-id="43bc9-137">displayNewMessageForm</span><span class="sxs-lookup"><span data-stu-id="43bc9-137">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="43bc9-138">メソッド</span><span class="sxs-lookup"><span data-stu-id="43bc9-138">Method</span></span> |
| [<span data-ttu-id="43bc9-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="43bc9-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="43bc9-140">メソッド</span><span class="sxs-lookup"><span data-stu-id="43bc9-140">Method</span></span> |
| [<span data-ttu-id="43bc9-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="43bc9-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="43bc9-142">メソッド</span><span class="sxs-lookup"><span data-stu-id="43bc9-142">Method</span></span> |
| [<span data-ttu-id="43bc9-143">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="43bc9-143">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="43bc9-144">メソッド</span><span class="sxs-lookup"><span data-stu-id="43bc9-144">Method</span></span> |
| [<span data-ttu-id="43bc9-145">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="43bc9-145">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="43bc9-146">メソッド</span><span class="sxs-lookup"><span data-stu-id="43bc9-146">Method</span></span> |
| [<span data-ttu-id="43bc9-147">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="43bc9-147">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="43bc9-148">メソッド</span><span class="sxs-lookup"><span data-stu-id="43bc9-148">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="43bc9-149">名前空間</span><span class="sxs-lookup"><span data-stu-id="43bc9-149">Namespaces</span></span>

<span data-ttu-id="43bc9-150">[diagnostics](Office.context.mailbox.diagnostics.md):Outlook アドインに診断情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="43bc9-150">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="43bc9-151">[item](Office.context.mailbox.item.md):Outlook アドインのメッセージや予定にアクセスするためのメソッドとプロパティを提供します。</span><span class="sxs-lookup"><span data-stu-id="43bc9-151">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="43bc9-152">[userProfile](Office.context.mailbox.userProfile.md):Outlook アドインのユーザーに関する情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="43bc9-152">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="43bc9-153">Members</span><span class="sxs-lookup"><span data-stu-id="43bc9-153">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="43bc9-154">ewsUrl: String</span><span class="sxs-lookup"><span data-stu-id="43bc9-154">ewsUrl: String</span></span>

<span data-ttu-id="43bc9-p101">このメール アカウントの Exchange Web サービス (EWS) エンドポイントの URL を取得します。読み取りモードのみです。</span><span class="sxs-lookup"><span data-stu-id="43bc9-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="43bc9-157">このメンバーは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="43bc9-157">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="43bc9-p102">
  `ewsUrl\` 値は、リモート サービスで、ユーザーのメールボックスに EWS 呼び出しを行うために使うことができます。たとえば、[選択したアイテムから添付ファイルを取得する](/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="43bc9-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="43bc9-160">アプリが閲覧モードで `ewsUrl` メンバーを呼び出すには、アプリのマニフェスト内に **ReadItem** アクセス許可が指定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="43bc9-160">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="43bc9-p103">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出してから、`ewsUrl` メンバーを使用する必要があります。アプリには、`saveAsync` メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="43bc9-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="43bc9-163">型</span><span class="sxs-lookup"><span data-stu-id="43bc9-163">Type</span></span>

*   <span data-ttu-id="43bc9-164">String</span><span class="sxs-lookup"><span data-stu-id="43bc9-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="43bc9-165">Requirements</span><span class="sxs-lookup"><span data-stu-id="43bc9-165">Requirements</span></span>

|<span data-ttu-id="43bc9-166">要件</span><span class="sxs-lookup"><span data-stu-id="43bc9-166">Requirement</span></span>| <span data-ttu-id="43bc9-167">値</span><span class="sxs-lookup"><span data-stu-id="43bc9-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="43bc9-168">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="43bc9-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43bc9-169">1.0</span><span class="sxs-lookup"><span data-stu-id="43bc9-169">1.0</span></span>|
|[<span data-ttu-id="43bc9-170">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="43bc9-170">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43bc9-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43bc9-171">ReadItem</span></span>|
|[<span data-ttu-id="43bc9-172">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="43bc9-172">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="43bc9-173">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="43bc9-173">Compose or Read</span></span>|

<br>

---
---

#### <a name="resturl-string"></a><span data-ttu-id="43bc9-174">restUrl: String</span><span class="sxs-lookup"><span data-stu-id="43bc9-174">restUrl: String</span></span>

<span data-ttu-id="43bc9-175">この電子メール アカウントの REST エンドポイントの URL を取得します。</span><span class="sxs-lookup"><span data-stu-id="43bc9-175">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="43bc9-176">`restUrl` 値は、ユーザーのメールボックスに [REST API](/outlook/rest/) 呼び出しを行うために使用できます。</span><span class="sxs-lookup"><span data-stu-id="43bc9-176">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="43bc9-177">型</span><span class="sxs-lookup"><span data-stu-id="43bc9-177">Type</span></span>

*   <span data-ttu-id="43bc9-178">String</span><span class="sxs-lookup"><span data-stu-id="43bc9-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="43bc9-179">Requirements</span><span class="sxs-lookup"><span data-stu-id="43bc9-179">Requirements</span></span>

|<span data-ttu-id="43bc9-180">要件</span><span class="sxs-lookup"><span data-stu-id="43bc9-180">Requirement</span></span>| <span data-ttu-id="43bc9-181">値</span><span class="sxs-lookup"><span data-stu-id="43bc9-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="43bc9-182">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="43bc9-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43bc9-183">1.5</span><span class="sxs-lookup"><span data-stu-id="43bc9-183">1.5</span></span> |
|[<span data-ttu-id="43bc9-184">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="43bc9-184">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43bc9-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43bc9-185">ReadItem</span></span>|
|[<span data-ttu-id="43bc9-186">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="43bc9-186">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="43bc9-187">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="43bc9-187">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="43bc9-188">メソッド</span><span class="sxs-lookup"><span data-stu-id="43bc9-188">Methods</span></span>

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="43bc9-189">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="43bc9-189">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="43bc9-190">サポートされているイベントのイベント ハンドラーを追加します。</span><span class="sxs-lookup"><span data-stu-id="43bc9-190">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="43bc9-191">現在サポートされている唯一のイベントの種類は `Office.EventType.ItemChanged` です。これはユーザーが新しいアイテムを選択したときに呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="43bc9-191">Currently the only supported event type is `Office.EventType.ItemChanged`, which is invoked when the user selects a new item.</span></span> <span data-ttu-id="43bc9-192">このイベントは、ピン留め可能な作業ウィンドウを実装するアドインで使用され、現在選択されているアイテムに基づいて作業ウィンドウ UI をアドインで更新できるようにします。</span><span class="sxs-lookup"><span data-stu-id="43bc9-192">This event is used by add-ins that implement a pinnable task pane, and allows the add-in to refresh the task pane UI based on the currently selected item.</span></span>

##### <a name="parameters"></a><span data-ttu-id="43bc9-193">パラメーター</span><span class="sxs-lookup"><span data-stu-id="43bc9-193">Parameters</span></span>

| <span data-ttu-id="43bc9-194">名前</span><span class="sxs-lookup"><span data-stu-id="43bc9-194">Name</span></span> | <span data-ttu-id="43bc9-195">種類</span><span class="sxs-lookup"><span data-stu-id="43bc9-195">Type</span></span> | <span data-ttu-id="43bc9-196">属性</span><span class="sxs-lookup"><span data-stu-id="43bc9-196">Attributes</span></span> | <span data-ttu-id="43bc9-197">説明</span><span class="sxs-lookup"><span data-stu-id="43bc9-197">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="43bc9-198">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="43bc9-198">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="43bc9-199">ハンドラーを呼び出す必要のあるイベント。</span><span class="sxs-lookup"><span data-stu-id="43bc9-199">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="43bc9-200">Function</span><span class="sxs-lookup"><span data-stu-id="43bc9-200">Function</span></span> || <span data-ttu-id="43bc9-p105">イベントを処理する関数。関数は、オブジェクト リテラルである単一パラメーターを受け入れる必要があります。パラメーターの `type` プロパティは、`addHandlerAsync` に渡される `eventType` パラメーターと一致します。</span><span class="sxs-lookup"><span data-stu-id="43bc9-p105">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="43bc9-204">Object</span><span class="sxs-lookup"><span data-stu-id="43bc9-204">Object</span></span> | <span data-ttu-id="43bc9-205">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="43bc9-205">&lt;optional&gt;</span></span> | <span data-ttu-id="43bc9-206">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="43bc9-206">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="43bc9-207">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="43bc9-207">Object</span></span> | <span data-ttu-id="43bc9-208">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="43bc9-208">&lt;optional&gt;</span></span> | <span data-ttu-id="43bc9-209">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="43bc9-209">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="43bc9-210">function</span><span class="sxs-lookup"><span data-stu-id="43bc9-210">function</span></span>| <span data-ttu-id="43bc9-211">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="43bc9-211">&lt;optional&gt;</span></span>|<span data-ttu-id="43bc9-212">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="43bc9-212">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="43bc9-213">Requirements</span><span class="sxs-lookup"><span data-stu-id="43bc9-213">Requirements</span></span>

|<span data-ttu-id="43bc9-214">要件</span><span class="sxs-lookup"><span data-stu-id="43bc9-214">Requirement</span></span>| <span data-ttu-id="43bc9-215">値</span><span class="sxs-lookup"><span data-stu-id="43bc9-215">Value</span></span>|
|---|---|
|[<span data-ttu-id="43bc9-216">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="43bc9-216">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43bc9-217">1.5</span><span class="sxs-lookup"><span data-stu-id="43bc9-217">1.5</span></span> |
|[<span data-ttu-id="43bc9-218">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="43bc9-218">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43bc9-219">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43bc9-219">ReadItem</span></span> |
|[<span data-ttu-id="43bc9-220">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="43bc9-220">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="43bc9-221">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="43bc9-221">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="43bc9-222">例</span><span class="sxs-lookup"><span data-stu-id="43bc9-222">Example</span></span>

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

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="43bc9-223">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="43bc9-223">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="43bc9-224">REST 形式のアイテム ID を EWS 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="43bc9-224">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="43bc9-225">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="43bc9-225">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="43bc9-p106">REST API ([Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) や [Microsoft Graph](https://graph.microsoft.io/) など) で取得されたアイテム ID は、Exchange Web サービス (EWS) に使用される形式とは異なる形式を使用します。`convertToEwsId` メソッドは、REST 形式の ID を EWS 用の適切な形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="43bc9-p106">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="43bc9-228">パラメーター</span><span class="sxs-lookup"><span data-stu-id="43bc9-228">Parameters</span></span>

|<span data-ttu-id="43bc9-229">名前</span><span class="sxs-lookup"><span data-stu-id="43bc9-229">Name</span></span>| <span data-ttu-id="43bc9-230">種類</span><span class="sxs-lookup"><span data-stu-id="43bc9-230">Type</span></span>| <span data-ttu-id="43bc9-231">説明</span><span class="sxs-lookup"><span data-stu-id="43bc9-231">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="43bc9-232">String</span><span class="sxs-lookup"><span data-stu-id="43bc9-232">String</span></span>|<span data-ttu-id="43bc9-233">Outlook REST API 形式のアイテム ID</span><span class="sxs-lookup"><span data-stu-id="43bc9-233">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="43bc9-234">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="43bc9-234">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.6)|<span data-ttu-id="43bc9-235">アイテム ID の取得に使用された Outlook REST API のバージョンを示す値。</span><span class="sxs-lookup"><span data-stu-id="43bc9-235">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="43bc9-236">要件</span><span class="sxs-lookup"><span data-stu-id="43bc9-236">Requirements</span></span>

|<span data-ttu-id="43bc9-237">要件</span><span class="sxs-lookup"><span data-stu-id="43bc9-237">Requirement</span></span>| <span data-ttu-id="43bc9-238">値</span><span class="sxs-lookup"><span data-stu-id="43bc9-238">Value</span></span>|
|---|---|
|[<span data-ttu-id="43bc9-239">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="43bc9-239">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43bc9-240">1.3</span><span class="sxs-lookup"><span data-stu-id="43bc9-240">1.3</span></span>|
|[<span data-ttu-id="43bc9-241">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="43bc9-241">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43bc9-242">制限あり</span><span class="sxs-lookup"><span data-stu-id="43bc9-242">Restricted</span></span>|
|[<span data-ttu-id="43bc9-243">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="43bc9-243">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="43bc9-244">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="43bc9-244">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="43bc9-245">戻り値:</span><span class="sxs-lookup"><span data-stu-id="43bc9-245">Returns:</span></span>

<span data-ttu-id="43bc9-246">型:String</span><span class="sxs-lookup"><span data-stu-id="43bc9-246">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="43bc9-247">例</span><span class="sxs-lookup"><span data-stu-id="43bc9-247">Example</span></span>

```js
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-16"></a><span data-ttu-id="43bc9-248">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.6)}</span><span class="sxs-lookup"><span data-stu-id="43bc9-248">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.6)}</span></span>

<span data-ttu-id="43bc9-249">クライアントのローカル時間で時間情報が含まれている辞書を取得します。</span><span class="sxs-lookup"><span data-stu-id="43bc9-249">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="43bc9-p107">Outlook on the web または Outlook デスクトップのメールアプリでは、日付と時刻に異なるタイムゾーンを使用できます。Outlook デスクトップは、クライアント コンピューターのタイム ゾーンを使用します。Outlook on the web は、Exchange 管理センター (EAC) で設定されたタイム ゾーンを使用します。日付と時刻の値は、ユーザー インターフェイスに表示される値が、常にユーザーが期待するタイム ゾーンと一致するように処理する必要があります。</span><span class="sxs-lookup"><span data-stu-id="43bc9-p107">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times. Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="43bc9-p108">Outlook デスクトップ クライアントでメール アプリを実行している場合、`convertToLocalClientTime` メソッドは、クライアント コンピューターのタイム ゾーンに設定された値のディクショナリ オブジェクトを返します。Outlook on the web でメール アプリを実行している場合、`convertToLocalClientTime` メソッドは、EAC で指定したタイム ゾーンに設定された値のディクショナリ オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="43bc9-p108">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="43bc9-255">パラメーター</span><span class="sxs-lookup"><span data-stu-id="43bc9-255">Parameters</span></span>

|<span data-ttu-id="43bc9-256">名前</span><span class="sxs-lookup"><span data-stu-id="43bc9-256">Name</span></span>| <span data-ttu-id="43bc9-257">種類</span><span class="sxs-lookup"><span data-stu-id="43bc9-257">Type</span></span>| <span data-ttu-id="43bc9-258">説明</span><span class="sxs-lookup"><span data-stu-id="43bc9-258">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="43bc9-259">日付</span><span class="sxs-lookup"><span data-stu-id="43bc9-259">Date</span></span>|<span data-ttu-id="43bc9-260">日付オブジェクト</span><span class="sxs-lookup"><span data-stu-id="43bc9-260">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="43bc9-261">Requirements</span><span class="sxs-lookup"><span data-stu-id="43bc9-261">Requirements</span></span>

|<span data-ttu-id="43bc9-262">要件</span><span class="sxs-lookup"><span data-stu-id="43bc9-262">Requirement</span></span>| <span data-ttu-id="43bc9-263">値</span><span class="sxs-lookup"><span data-stu-id="43bc9-263">Value</span></span>|
|---|---|
|[<span data-ttu-id="43bc9-264">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="43bc9-264">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43bc9-265">1.0</span><span class="sxs-lookup"><span data-stu-id="43bc9-265">1.0</span></span>|
|[<span data-ttu-id="43bc9-266">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="43bc9-266">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43bc9-267">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43bc9-267">ReadItem</span></span>|
|[<span data-ttu-id="43bc9-268">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="43bc9-268">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="43bc9-269">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="43bc9-269">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="43bc9-270">戻り値:</span><span class="sxs-lookup"><span data-stu-id="43bc9-270">Returns:</span></span>

<span data-ttu-id="43bc9-271">型:[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="43bc9-271">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.6)</span></span>

<br>

---
---

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="43bc9-272">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="43bc9-272">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="43bc9-273">EWS 形式のアイテム ID を REST 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="43bc9-273">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="43bc9-274">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="43bc9-274">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="43bc9-p109">EWS または `itemId` プロパティで取得されるアイテム ID は、REST API ([Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) や [Microsoft Graph](https://graph.microsoft.io/) など) に使用される形式とは異なる形式を使用します。`convertToRestId` メソッドは、EWS 形式の ID を REST 用の適切な形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="43bc9-p109">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="43bc9-277">パラメーター</span><span class="sxs-lookup"><span data-stu-id="43bc9-277">Parameters</span></span>

|<span data-ttu-id="43bc9-278">名前</span><span class="sxs-lookup"><span data-stu-id="43bc9-278">Name</span></span>| <span data-ttu-id="43bc9-279">型</span><span class="sxs-lookup"><span data-stu-id="43bc9-279">Type</span></span>| <span data-ttu-id="43bc9-280">説明</span><span class="sxs-lookup"><span data-stu-id="43bc9-280">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="43bc9-281">String</span><span class="sxs-lookup"><span data-stu-id="43bc9-281">String</span></span>|<span data-ttu-id="43bc9-282">Exchange Web サービス (EWS) 形式のアイテム ID</span><span class="sxs-lookup"><span data-stu-id="43bc9-282">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="43bc9-283">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="43bc9-283">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.6)|<span data-ttu-id="43bc9-284">変換後の ID を使用する Outlook REST API のバージョンを示す値。</span><span class="sxs-lookup"><span data-stu-id="43bc9-284">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="43bc9-285">Requirements</span><span class="sxs-lookup"><span data-stu-id="43bc9-285">Requirements</span></span>

|<span data-ttu-id="43bc9-286">要件</span><span class="sxs-lookup"><span data-stu-id="43bc9-286">Requirement</span></span>| <span data-ttu-id="43bc9-287">値</span><span class="sxs-lookup"><span data-stu-id="43bc9-287">Value</span></span>|
|---|---|
|[<span data-ttu-id="43bc9-288">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="43bc9-288">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43bc9-289">1.3</span><span class="sxs-lookup"><span data-stu-id="43bc9-289">1.3</span></span>|
|[<span data-ttu-id="43bc9-290">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="43bc9-290">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43bc9-291">制限あり</span><span class="sxs-lookup"><span data-stu-id="43bc9-291">Restricted</span></span>|
|[<span data-ttu-id="43bc9-292">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="43bc9-292">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="43bc9-293">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="43bc9-293">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="43bc9-294">戻り値:</span><span class="sxs-lookup"><span data-stu-id="43bc9-294">Returns:</span></span>

<span data-ttu-id="43bc9-295">型:String</span><span class="sxs-lookup"><span data-stu-id="43bc9-295">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="43bc9-296">例</span><span class="sxs-lookup"><span data-stu-id="43bc9-296">Example</span></span>

```js
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="43bc9-297">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="43bc9-297">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="43bc9-298">時間情報が含まれているディクショナリから日付オブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="43bc9-298">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="43bc9-299">`convertToUtcClientTime` メソッドは、ローカルの日付と時刻を含むディクショナリを、ローカルの日付と時刻の正しい値を持つ日付オブジェクトに変換します。</span><span class="sxs-lookup"><span data-stu-id="43bc9-299">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="43bc9-300">パラメーター</span><span class="sxs-lookup"><span data-stu-id="43bc9-300">Parameters</span></span>

|<span data-ttu-id="43bc9-301">名前</span><span class="sxs-lookup"><span data-stu-id="43bc9-301">Name</span></span>| <span data-ttu-id="43bc9-302">型</span><span class="sxs-lookup"><span data-stu-id="43bc9-302">Type</span></span>| <span data-ttu-id="43bc9-303">説明</span><span class="sxs-lookup"><span data-stu-id="43bc9-303">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="43bc9-304">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="43bc9-304">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.6)|<span data-ttu-id="43bc9-305">変換するローカル時刻の値。</span><span class="sxs-lookup"><span data-stu-id="43bc9-305">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="43bc9-306">Requirements</span><span class="sxs-lookup"><span data-stu-id="43bc9-306">Requirements</span></span>

|<span data-ttu-id="43bc9-307">要件</span><span class="sxs-lookup"><span data-stu-id="43bc9-307">Requirement</span></span>| <span data-ttu-id="43bc9-308">値</span><span class="sxs-lookup"><span data-stu-id="43bc9-308">Value</span></span>|
|---|---|
|[<span data-ttu-id="43bc9-309">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="43bc9-309">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43bc9-310">1.0</span><span class="sxs-lookup"><span data-stu-id="43bc9-310">1.0</span></span>|
|[<span data-ttu-id="43bc9-311">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="43bc9-311">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43bc9-312">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43bc9-312">ReadItem</span></span>|
|[<span data-ttu-id="43bc9-313">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="43bc9-313">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="43bc9-314">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="43bc9-314">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="43bc9-315">戻り値:</span><span class="sxs-lookup"><span data-stu-id="43bc9-315">Returns:</span></span>

<span data-ttu-id="43bc9-316">時間が UTC で表現された日付オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="43bc9-316">A Date object with the time expressed in UTC.</span></span>

<span data-ttu-id="43bc9-317">型: Date</span><span class="sxs-lookup"><span data-stu-id="43bc9-317">Type: Date</span></span>

##### <a name="example"></a><span data-ttu-id="43bc9-318">例</span><span class="sxs-lookup"><span data-stu-id="43bc9-318">Example</span></span>

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

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="43bc9-319">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="43bc9-319">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="43bc9-320">既存の予定を表示します。</span><span class="sxs-lookup"><span data-stu-id="43bc9-320">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="43bc9-321">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="43bc9-321">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="43bc9-322">`displayAppointmentForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存の予定を開きます。</span><span class="sxs-lookup"><span data-stu-id="43bc9-322">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="43bc9-p110">Outlook on Mac では、このメソッドを使用して定期的な系列に含まれない単発の予定や定期的な系列のマスター予定を表示できます。ただし、系列のインスタンスは表示できません。これは、Outlook on Mac では定期的な系列のインスタンスのプロパティ (アイテム ID を含む) にアクセスできないためです。</span><span class="sxs-lookup"><span data-stu-id="43bc9-p110">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="43bc9-325">Outlook on the web では、このメソッドはフォームの本文が 32KB 以下の文字数の場合にのみ指定のフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="43bc9-325">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="43bc9-326">指定のアイテム識別子が既存の予定を表していない場合は、クライアント コンピューターまたはデバイスで空のウィンドウが開き、エラー メッセージは返されません。</span><span class="sxs-lookup"><span data-stu-id="43bc9-326">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="43bc9-327">パラメーター</span><span class="sxs-lookup"><span data-stu-id="43bc9-327">Parameters</span></span>

|<span data-ttu-id="43bc9-328">名前</span><span class="sxs-lookup"><span data-stu-id="43bc9-328">Name</span></span>| <span data-ttu-id="43bc9-329">型</span><span class="sxs-lookup"><span data-stu-id="43bc9-329">Type</span></span>| <span data-ttu-id="43bc9-330">説明</span><span class="sxs-lookup"><span data-stu-id="43bc9-330">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="43bc9-331">String</span><span class="sxs-lookup"><span data-stu-id="43bc9-331">String</span></span>|<span data-ttu-id="43bc9-332">既存の予定の Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="43bc9-332">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="43bc9-333">Requirements</span><span class="sxs-lookup"><span data-stu-id="43bc9-333">Requirements</span></span>

|<span data-ttu-id="43bc9-334">要件</span><span class="sxs-lookup"><span data-stu-id="43bc9-334">Requirement</span></span>| <span data-ttu-id="43bc9-335">値</span><span class="sxs-lookup"><span data-stu-id="43bc9-335">Value</span></span>|
|---|---|
|[<span data-ttu-id="43bc9-336">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="43bc9-336">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43bc9-337">1.0</span><span class="sxs-lookup"><span data-stu-id="43bc9-337">1.0</span></span>|
|[<span data-ttu-id="43bc9-338">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="43bc9-338">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43bc9-339">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43bc9-339">ReadItem</span></span>|
|[<span data-ttu-id="43bc9-340">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="43bc9-340">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="43bc9-341">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="43bc9-341">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="43bc9-342">例</span><span class="sxs-lookup"><span data-stu-id="43bc9-342">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

<br>

---
---

#### <a name="displaymessageformitemid"></a><span data-ttu-id="43bc9-343">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="43bc9-343">displayMessageForm(itemId)</span></span>

<span data-ttu-id="43bc9-344">既存のメッセージを表示します。</span><span class="sxs-lookup"><span data-stu-id="43bc9-344">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="43bc9-345">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="43bc9-345">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="43bc9-346">`displayMessageForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存のメッセージを開きます。</span><span class="sxs-lookup"><span data-stu-id="43bc9-346">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="43bc9-347">Outlook on the web では、このメソッドはフォームの本文が 32 KB 以下の文字数の場合にのみ指定のフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="43bc9-347">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="43bc9-348">指定のアイテム識別子が既存のメッセージを表していない場合は、クライアント コンピューターにはメッセージは表示されず、エラー メッセージも返されません。</span><span class="sxs-lookup"><span data-stu-id="43bc9-348">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="43bc9-p111">予定を表す `itemId` を含む `displayMessageForm` を使用しないでください。既存の予定を表示するには、`displayAppointmentForm` メソッドを使用します。新しい予定を作成するフォームを表示するには、`displayNewAppointmentForm` メソッドを使用します。</span><span class="sxs-lookup"><span data-stu-id="43bc9-p111">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="43bc9-351">パラメーター</span><span class="sxs-lookup"><span data-stu-id="43bc9-351">Parameters</span></span>

|<span data-ttu-id="43bc9-352">名前</span><span class="sxs-lookup"><span data-stu-id="43bc9-352">Name</span></span>| <span data-ttu-id="43bc9-353">型</span><span class="sxs-lookup"><span data-stu-id="43bc9-353">Type</span></span>| <span data-ttu-id="43bc9-354">説明</span><span class="sxs-lookup"><span data-stu-id="43bc9-354">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="43bc9-355">String</span><span class="sxs-lookup"><span data-stu-id="43bc9-355">String</span></span>|<span data-ttu-id="43bc9-356">既存のメッセージの Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="43bc9-356">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="43bc9-357">要件</span><span class="sxs-lookup"><span data-stu-id="43bc9-357">Requirements</span></span>

|<span data-ttu-id="43bc9-358">要件</span><span class="sxs-lookup"><span data-stu-id="43bc9-358">Requirement</span></span>| <span data-ttu-id="43bc9-359">値</span><span class="sxs-lookup"><span data-stu-id="43bc9-359">Value</span></span>|
|---|---|
|[<span data-ttu-id="43bc9-360">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="43bc9-360">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43bc9-361">1.0</span><span class="sxs-lookup"><span data-stu-id="43bc9-361">1.0</span></span>|
|[<span data-ttu-id="43bc9-362">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="43bc9-362">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43bc9-363">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43bc9-363">ReadItem</span></span>|
|[<span data-ttu-id="43bc9-364">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="43bc9-364">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="43bc9-365">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="43bc9-365">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="43bc9-366">例</span><span class="sxs-lookup"><span data-stu-id="43bc9-366">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

<br>

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="43bc9-367">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="43bc9-367">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="43bc9-368">新しい予定を作成するためのフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="43bc9-368">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="43bc9-369">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="43bc9-369">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="43bc9-p112">`displayNewAppointmentForm` メソッドを使用すると、ユーザーが新しい予定または会議を作成できるフォームが開きます。パラメーターを指定すると、予定のフォーム フィールドにパラメーターの内容が自動的に設定されます。</span><span class="sxs-lookup"><span data-stu-id="43bc9-p112">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="43bc9-p113">Outlook on the web およびモバイル デバイスでは、このメソッドは常に出席者フィールドが含まれるフォームを表示します。入力引数として出席者を指定しないと、このメソッドは **[保存]** ボタンのあるフォームを表示します。出席者を指定した場合には、フォームにその出席者と **[送信]** ボタンが表示されます。</span><span class="sxs-lookup"><span data-stu-id="43bc9-p113">In Outlook on the web and mobile devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="43bc9-p114">Outlook リッチ クライアントと Outlook RT で、`requiredAttendees`、`optionalAttendees`、または `resources` パラメーターに出席者またはリソースを指定し、このメソッドを実行すると、**[送信]** ボタンがある会議フォームが表示されます。受信者を指定せずにこのメソッドを実行すると、**[保存して閉じる]** ボタンがある予定フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="43bc9-p114">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="43bc9-377">パラメーターのいずれかが指定のサイズ制限を超える場合、または不明なパラメーター名が指定されている場合は、例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="43bc9-377">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="43bc9-378">パラメーター</span><span class="sxs-lookup"><span data-stu-id="43bc9-378">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="43bc9-379">すべてのパラメーターは省略可能です。</span><span class="sxs-lookup"><span data-stu-id="43bc9-379">All parameters are optional.</span></span>

|<span data-ttu-id="43bc9-380">名前</span><span class="sxs-lookup"><span data-stu-id="43bc9-380">Name</span></span>| <span data-ttu-id="43bc9-381">種類</span><span class="sxs-lookup"><span data-stu-id="43bc9-381">Type</span></span>| <span data-ttu-id="43bc9-382">説明</span><span class="sxs-lookup"><span data-stu-id="43bc9-382">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="43bc9-383">Object</span><span class="sxs-lookup"><span data-stu-id="43bc9-383">Object</span></span> | <span data-ttu-id="43bc9-384">新しい予定を記述するパラメーターのディクショナリ。</span><span class="sxs-lookup"><span data-stu-id="43bc9-384">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="43bc9-385">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span><span class="sxs-lookup"><span data-stu-id="43bc9-385">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span></span> | <span data-ttu-id="43bc9-p115">予定に必要な各出席者について、メール アドレスを含む文字列の配列、または `EmailAddressDetails` オブジェクトを含む配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="43bc9-p115">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="43bc9-388">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span><span class="sxs-lookup"><span data-stu-id="43bc9-388">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span></span> | <span data-ttu-id="43bc9-p116">予定の各任意出席者について、メール アドレスを含む文字列の配列、または `EmailAddressDetails` オブジェクトを含む配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="43bc9-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="43bc9-391">日付</span><span class="sxs-lookup"><span data-stu-id="43bc9-391">Date</span></span> | <span data-ttu-id="43bc9-392">予定の開始日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="43bc9-392">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="43bc9-393">日付</span><span class="sxs-lookup"><span data-stu-id="43bc9-393">Date</span></span> | <span data-ttu-id="43bc9-394">予定の終了日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="43bc9-394">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="43bc9-395">String</span><span class="sxs-lookup"><span data-stu-id="43bc9-395">String</span></span> | <span data-ttu-id="43bc9-p117">予定の場所を含む文字列。文字列は最大 255 文字に制限されます。</span><span class="sxs-lookup"><span data-stu-id="43bc9-p117">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="43bc9-398">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="43bc9-398">Array.&lt;String&gt;</span></span> | <span data-ttu-id="43bc9-p118">予定に必要なリソースを含む文字列の配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="43bc9-p118">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="43bc9-401">String</span><span class="sxs-lookup"><span data-stu-id="43bc9-401">String</span></span> | <span data-ttu-id="43bc9-p119">予定の件名を含む文字列です。文字列は最大 255 文字に制限されます。</span><span class="sxs-lookup"><span data-stu-id="43bc9-p119">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="43bc9-404">String</span><span class="sxs-lookup"><span data-stu-id="43bc9-404">String</span></span> | <span data-ttu-id="43bc9-p120">予定の本文。本文の内容は、最大サイズが 32 KB に制限されます。</span><span class="sxs-lookup"><span data-stu-id="43bc9-p120">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="43bc9-407">要件</span><span class="sxs-lookup"><span data-stu-id="43bc9-407">Requirements</span></span>

|<span data-ttu-id="43bc9-408">要件</span><span class="sxs-lookup"><span data-stu-id="43bc9-408">Requirement</span></span>| <span data-ttu-id="43bc9-409">値</span><span class="sxs-lookup"><span data-stu-id="43bc9-409">Value</span></span>|
|---|---|
|[<span data-ttu-id="43bc9-410">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="43bc9-410">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43bc9-411">1.0</span><span class="sxs-lookup"><span data-stu-id="43bc9-411">1.0</span></span>|
|[<span data-ttu-id="43bc9-412">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="43bc9-412">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43bc9-413">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43bc9-413">ReadItem</span></span>|
|[<span data-ttu-id="43bc9-414">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="43bc9-414">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="43bc9-415">読み取り</span><span class="sxs-lookup"><span data-stu-id="43bc9-415">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="43bc9-416">例</span><span class="sxs-lookup"><span data-stu-id="43bc9-416">Example</span></span>

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

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="43bc9-417">displayNewMessageForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="43bc9-417">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="43bc9-418">新しいメッセージを作成するためのフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="43bc9-418">Displays a form for creating a new message.</span></span>

<span data-ttu-id="43bc9-p121">The `displayNewMessageForm` method opens a form that enables the user to create a new message. If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span><span class="sxs-lookup"><span data-stu-id="43bc9-p121">The `displayNewMessageForm` method opens a form that enables the user to create a new message. If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="43bc9-421">パラメーターのいずれかが指定のサイズ制限を超える場合、または不明なパラメーター名が指定されている場合は、例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="43bc9-421">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="43bc9-422">パラメーター</span><span class="sxs-lookup"><span data-stu-id="43bc9-422">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="43bc9-423">すべてのパラメーターは省略可能です。</span><span class="sxs-lookup"><span data-stu-id="43bc9-423">All parameters are optional.</span></span>

|<span data-ttu-id="43bc9-424">名前</span><span class="sxs-lookup"><span data-stu-id="43bc9-424">Name</span></span>| <span data-ttu-id="43bc9-425">種類</span><span class="sxs-lookup"><span data-stu-id="43bc9-425">Type</span></span>| <span data-ttu-id="43bc9-426">説明</span><span class="sxs-lookup"><span data-stu-id="43bc9-426">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="43bc9-427">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="43bc9-427">Object</span></span> | <span data-ttu-id="43bc9-428">新しいメッセージを記述するパラメータの辞書。</span><span class="sxs-lookup"><span data-stu-id="43bc9-428">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="43bc9-429">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span><span class="sxs-lookup"><span data-stu-id="43bc9-429">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span></span> | <span data-ttu-id="43bc9-p122">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line. The array is limited to a maximum of 100 entries.</span><span class="sxs-lookup"><span data-stu-id="43bc9-p122">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="43bc9-432">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span><span class="sxs-lookup"><span data-stu-id="43bc9-432">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span></span> | <span data-ttu-id="43bc9-p123">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line. The array is limited to a maximum of 100 entries.</span><span class="sxs-lookup"><span data-stu-id="43bc9-p123">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="43bc9-435">配列。&lt;文字列&gt; | 配列。&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span><span class="sxs-lookup"><span data-stu-id="43bc9-435">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span></span> | <span data-ttu-id="43bc9-p124">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line. The array is limited to a maximum of 100 entries.</span><span class="sxs-lookup"><span data-stu-id="43bc9-p124">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="43bc9-438">String</span><span class="sxs-lookup"><span data-stu-id="43bc9-438">String</span></span> | <span data-ttu-id="43bc9-p125">A string containing the subject of the message. The string is limited to a maximum of 255 characters.</span><span class="sxs-lookup"><span data-stu-id="43bc9-p125">A string containing the subject of the message. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="43bc9-441">String</span><span class="sxs-lookup"><span data-stu-id="43bc9-441">String</span></span> | <span data-ttu-id="43bc9-p126">The HTML body of the message. The body content is limited to a maximum size of 32 KB.</span><span class="sxs-lookup"><span data-stu-id="43bc9-p126">The HTML body of the message. The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="43bc9-444">配列。&lt;オブジェクト&gt;</span><span class="sxs-lookup"><span data-stu-id="43bc9-444">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="43bc9-445">添付ファイルまたは添付アイテムである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="43bc9-445">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="43bc9-446">String</span><span class="sxs-lookup"><span data-stu-id="43bc9-446">String</span></span> | <span data-ttu-id="43bc9-p127">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="43bc9-p127">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="43bc9-449">String</span><span class="sxs-lookup"><span data-stu-id="43bc9-449">String</span></span> | <span data-ttu-id="43bc9-450">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="43bc9-450">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="43bc9-451">文字列</span><span class="sxs-lookup"><span data-stu-id="43bc9-451">String</span></span> | <span data-ttu-id="43bc9-p128">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="43bc9-p128">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="43bc9-454">ブール値</span><span class="sxs-lookup"><span data-stu-id="43bc9-454">Boolean</span></span> | <span data-ttu-id="43bc9-p129">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="43bc9-p129">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="43bc9-457">String</span><span class="sxs-lookup"><span data-stu-id="43bc9-457">String</span></span> | <span data-ttu-id="43bc9-p130">Only used if `type` is set to `item`. The EWS item id of the existing e-mail you want to attach to the new message. This is a string up to 100 characters.</span><span class="sxs-lookup"><span data-stu-id="43bc9-p130">Only used if `type` is set to `item`. The EWS item id of the existing e-mail you want to attach to the new message. This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="43bc9-461">Requirements</span><span class="sxs-lookup"><span data-stu-id="43bc9-461">Requirements</span></span>

|<span data-ttu-id="43bc9-462">要件</span><span class="sxs-lookup"><span data-stu-id="43bc9-462">Requirement</span></span>| <span data-ttu-id="43bc9-463">値</span><span class="sxs-lookup"><span data-stu-id="43bc9-463">Value</span></span>|
|---|---|
|[<span data-ttu-id="43bc9-464">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="43bc9-464">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43bc9-465">1.6</span><span class="sxs-lookup"><span data-stu-id="43bc9-465">1.6</span></span> |
|[<span data-ttu-id="43bc9-466">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="43bc9-466">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43bc9-467">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43bc9-467">ReadItem</span></span>|
|[<span data-ttu-id="43bc9-468">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="43bc9-468">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="43bc9-469">読み取り</span><span class="sxs-lookup"><span data-stu-id="43bc9-469">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="43bc9-470">例</span><span class="sxs-lookup"><span data-stu-id="43bc9-470">Example</span></span>

```js
Office.context.mailbox.displayNewMessageForm(
  {
    toRecipients: Office.context.mailbox.item.to, // Copy the To line from current item
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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="43bc9-471">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="43bc9-471">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="43bc9-472">REST API または Exchange Web サービスを呼び出すために使用するトークンを含む文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="43bc9-472">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="43bc9-p131">`getCallbackTokenAsync` メソッドは、ユーザーのメールボックスをホストする Exchange Server から不透明なトークンを取得する非同期の呼び出しを行います。コールバック トークンの有効期間は 5 分です。</span><span class="sxs-lookup"><span data-stu-id="43bc9-p131">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="43bc9-475">可能であれば、アドインでは Exchange Web サービスの代わりに REST API を使用することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="43bc9-475">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span>

<span data-ttu-id="43bc9-476">閲覧モードで `getCallbackTokenAsync` メソッドを呼び出すには、**ReadItem** の最小限のアクセス許可レベルが必要です。</span><span class="sxs-lookup"><span data-stu-id="43bc9-476">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="43bc9-477">作成モードで `getCallbackTokenAsync` を呼び出すには、アイテムを保存しておく必要があります。</span><span class="sxs-lookup"><span data-stu-id="43bc9-477">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="43bc9-478">[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドは、**ReadWriteItem** の最小限のアクセス許可レベルが必要です。</span><span class="sxs-lookup"><span data-stu-id="43bc9-478">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

<span data-ttu-id="43bc9-479">**REST トークン**</span><span class="sxs-lookup"><span data-stu-id="43bc9-479">**REST Tokens**</span></span>

<span data-ttu-id="43bc9-p133">REST トークンが要求された場合 (`options.isRest = true`)、結果トークンは Exchange Web サービスの呼び出しを認証するためには機能しません。アドインがマニフェストで [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) アクセス許可を指定していない限り、トークンの範囲は現在のアイテムとその添付ファイルへの読み取り専用アクセスに制限されます。`ReadWriteMailbox` アクセス許可が指定されている場合は、結果トークンは、メールを送信する機能など、メール、カレンダー、連絡先への読み取り/書き込みアクセスを付与します。</span><span class="sxs-lookup"><span data-stu-id="43bc9-p133">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="43bc9-483">アドインでは、`restUrl` プロパティを使用して、REST API 呼び出しを行うときに使用する正しい URL を決定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="43bc9-483">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="43bc9-484">**EWS トークン**</span><span class="sxs-lookup"><span data-stu-id="43bc9-484">**EWS Tokens**</span></span>

<span data-ttu-id="43bc9-p134">EWS トークンが要求された場合 (`options.isRest = false`)、結果トークンは REST API 呼び出しを認証するためには機能しません。トークンの範囲は、現在のアイテムへのアクセスに制限されます。</span><span class="sxs-lookup"><span data-stu-id="43bc9-p134">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="43bc9-487">アドインでは、`ewsUrl` プロパティを使用して、EWS 呼び出しを行うときに使用する正しい URL を決定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="43bc9-487">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

<span data-ttu-id="43bc9-488">トークンと、添付ファイル識別子またはアイテム識別子の両方をサードパーティ システムに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="43bc9-488">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="43bc9-489">サードパーティ システムは、トークンをベアラー承認トークンとして使用し、Exchange Web サービス (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) 操作または [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) 操作を呼び出して、添付ファイルまたはアイテムを返します。</span><span class="sxs-lookup"><span data-stu-id="43bc9-489">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item.</span></span> <span data-ttu-id="43bc9-490">たとえば、[選択したアイテムから添付ファイルを取得する](/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="43bc9-490">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

##### <a name="parameters"></a><span data-ttu-id="43bc9-491">パラメーター</span><span class="sxs-lookup"><span data-stu-id="43bc9-491">Parameters</span></span>

|<span data-ttu-id="43bc9-492">名前</span><span class="sxs-lookup"><span data-stu-id="43bc9-492">Name</span></span>| <span data-ttu-id="43bc9-493">型</span><span class="sxs-lookup"><span data-stu-id="43bc9-493">Type</span></span>| <span data-ttu-id="43bc9-494">属性</span><span class="sxs-lookup"><span data-stu-id="43bc9-494">Attributes</span></span>| <span data-ttu-id="43bc9-495">説明</span><span class="sxs-lookup"><span data-stu-id="43bc9-495">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="43bc9-496">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="43bc9-496">Object</span></span> | <span data-ttu-id="43bc9-497">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="43bc9-497">&lt;optional&gt;</span></span> | <span data-ttu-id="43bc9-498">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="43bc9-498">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="43bc9-499">ブール値</span><span class="sxs-lookup"><span data-stu-id="43bc9-499">Boolean</span></span> |  <span data-ttu-id="43bc9-500">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="43bc9-500">&lt;optional&gt;</span></span> | <span data-ttu-id="43bc9-p136">提供されたトークンを Outlook REST API または Exchange Web サービスに使用するかどうかを決定します。既定値は、`false` です。</span><span class="sxs-lookup"><span data-stu-id="43bc9-p136">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="43bc9-503">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="43bc9-503">Object</span></span> |  <span data-ttu-id="43bc9-504">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="43bc9-504">&lt;optional&gt;</span></span> | <span data-ttu-id="43bc9-505">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="43bc9-505">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="43bc9-506">function</span><span class="sxs-lookup"><span data-stu-id="43bc9-506">function</span></span>||<span data-ttu-id="43bc9-507">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="43bc9-507">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="43bc9-508">トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="43bc9-508">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="43bc9-509">エラーが発生した場合、 `asyncResult.error` および `asyncResult.diagnostics` のプロパティで追加情報が提供される場合があります。</span><span class="sxs-lookup"><span data-stu-id="43bc9-509">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|

##### <a name="errors"></a><span data-ttu-id="43bc9-510">エラー</span><span class="sxs-lookup"><span data-stu-id="43bc9-510">Errors</span></span>

|<span data-ttu-id="43bc9-511">エラー コード</span><span class="sxs-lookup"><span data-stu-id="43bc9-511">Error code</span></span>|<span data-ttu-id="43bc9-512">説明</span><span class="sxs-lookup"><span data-stu-id="43bc9-512">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="43bc9-513">要求が失敗しました。</span><span class="sxs-lookup"><span data-stu-id="43bc9-513">The request has failed.</span></span> <span data-ttu-id="43bc9-514">HTTP エラーコードの diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="43bc9-514">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="43bc9-515">Exchange サーバーがエラーを返しました。</span><span class="sxs-lookup"><span data-stu-id="43bc9-515">The Exchange server returned an error.</span></span> <span data-ttu-id="43bc9-516">詳細については、diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="43bc9-516">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="43bc9-517">ユーザーはネットワークに接続されていません。</span><span class="sxs-lookup"><span data-stu-id="43bc9-517">The user is no longer connected to the network.</span></span> <span data-ttu-id="43bc9-518">ネットワーク接続を確認し、やり直してください。</span><span class="sxs-lookup"><span data-stu-id="43bc9-518">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="43bc9-519">要件</span><span class="sxs-lookup"><span data-stu-id="43bc9-519">Requirements</span></span>

|<span data-ttu-id="43bc9-520">必要条件</span><span class="sxs-lookup"><span data-stu-id="43bc9-520">Requirement</span></span>| <span data-ttu-id="43bc9-521">値</span><span class="sxs-lookup"><span data-stu-id="43bc9-521">Value</span></span>|
|---|---|
|[<span data-ttu-id="43bc9-522">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="43bc9-522">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43bc9-523">1.5</span><span class="sxs-lookup"><span data-stu-id="43bc9-523">1.5</span></span> |
|[<span data-ttu-id="43bc9-524">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="43bc9-524">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43bc9-525">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43bc9-525">ReadItem</span></span>|
|[<span data-ttu-id="43bc9-526">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="43bc9-526">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="43bc9-527">新規作成と閲覧</span><span class="sxs-lookup"><span data-stu-id="43bc9-527">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="43bc9-528">例</span><span class="sxs-lookup"><span data-stu-id="43bc9-528">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="43bc9-529">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="43bc9-529">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="43bc9-530">Exchange Server から添付ファイルやアイテムを取得するために使うトークンを含む文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="43bc9-530">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="43bc9-p140">`getCallbackTokenAsync` メソッドは、ユーザーのメールボックスをホストする Exchange Server から不透明なトークンを取得する非同期の呼び出しを行います。コールバック トークンの有効期間は 5 分です。</span><span class="sxs-lookup"><span data-stu-id="43bc9-p140">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="43bc9-533">トークンと、添付ファイル識別子またはアイテム識別子の両方をサードパーティ システムに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="43bc9-533">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="43bc9-534">サードパーティ システムは、トークンをベアラー承認トークンとして使用し、Exchange Web サービス (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) 操作または [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) 操作を呼び出して、添付ファイルまたはアイテムを返します。</span><span class="sxs-lookup"><span data-stu-id="43bc9-534">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item.</span></span> <span data-ttu-id="43bc9-535">たとえば、[選択したアイテムから添付ファイルを取得する](/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="43bc9-535">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="43bc9-536">閲覧モードで `getCallbackTokenAsync` メソッドを呼び出すには、**ReadItem** の最小限のアクセス許可レベルが必要です。</span><span class="sxs-lookup"><span data-stu-id="43bc9-536">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="43bc9-537">作成モードで `getCallbackTokenAsync` を呼び出すには、アイテムを保存しておく必要があります。</span><span class="sxs-lookup"><span data-stu-id="43bc9-537">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="43bc9-538">[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドは、**ReadWriteItem** の最小限のアクセス許可レベルが必要です。</span><span class="sxs-lookup"><span data-stu-id="43bc9-538">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

##### <a name="parameters"></a><span data-ttu-id="43bc9-539">パラメーター</span><span class="sxs-lookup"><span data-stu-id="43bc9-539">Parameters</span></span>

|<span data-ttu-id="43bc9-540">名前</span><span class="sxs-lookup"><span data-stu-id="43bc9-540">Name</span></span>| <span data-ttu-id="43bc9-541">型</span><span class="sxs-lookup"><span data-stu-id="43bc9-541">Type</span></span>| <span data-ttu-id="43bc9-542">属性</span><span class="sxs-lookup"><span data-stu-id="43bc9-542">Attributes</span></span>| <span data-ttu-id="43bc9-543">説明</span><span class="sxs-lookup"><span data-stu-id="43bc9-543">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="43bc9-544">function</span><span class="sxs-lookup"><span data-stu-id="43bc9-544">function</span></span>||<span data-ttu-id="43bc9-545">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="43bc9-545">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="43bc9-546">トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="43bc9-546">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="43bc9-547">エラーが発生した場合、 `asyncResult.error` および `asyncResult.diagnostics` のプロパティで追加情報が提供される場合があります。</span><span class="sxs-lookup"><span data-stu-id="43bc9-547">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="43bc9-548">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="43bc9-548">Object</span></span>| <span data-ttu-id="43bc9-549">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="43bc9-549">&lt;optional&gt;</span></span>|<span data-ttu-id="43bc9-550">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="43bc9-550">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="43bc9-551">エラー</span><span class="sxs-lookup"><span data-stu-id="43bc9-551">Errors</span></span>

|<span data-ttu-id="43bc9-552">エラー コード</span><span class="sxs-lookup"><span data-stu-id="43bc9-552">Error code</span></span>|<span data-ttu-id="43bc9-553">説明</span><span class="sxs-lookup"><span data-stu-id="43bc9-553">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="43bc9-554">要求が失敗しました。</span><span class="sxs-lookup"><span data-stu-id="43bc9-554">The request has failed.</span></span> <span data-ttu-id="43bc9-555">HTTP エラーコードの diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="43bc9-555">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="43bc9-556">Exchange サーバーがエラーを返しました。</span><span class="sxs-lookup"><span data-stu-id="43bc9-556">The Exchange server returned an error.</span></span> <span data-ttu-id="43bc9-557">詳細については、diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="43bc9-557">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="43bc9-558">ユーザーはネットワークに接続されていません。</span><span class="sxs-lookup"><span data-stu-id="43bc9-558">The user is no longer connected to the network.</span></span> <span data-ttu-id="43bc9-559">ネットワーク接続を確認し、やり直してください。</span><span class="sxs-lookup"><span data-stu-id="43bc9-559">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="43bc9-560">要件</span><span class="sxs-lookup"><span data-stu-id="43bc9-560">Requirements</span></span>

|<span data-ttu-id="43bc9-561">要件</span><span class="sxs-lookup"><span data-stu-id="43bc9-561">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="43bc9-562">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="43bc9-562">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43bc9-563">1.0</span><span class="sxs-lookup"><span data-stu-id="43bc9-563">1.0</span></span> | <span data-ttu-id="43bc9-564">1.3</span><span class="sxs-lookup"><span data-stu-id="43bc9-564">1.3</span></span> |
|[<span data-ttu-id="43bc9-565">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="43bc9-565">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43bc9-566">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43bc9-566">ReadItem</span></span> | <span data-ttu-id="43bc9-567">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43bc9-567">ReadItem</span></span> |
|[<span data-ttu-id="43bc9-568">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="43bc9-568">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="43bc9-569">Read</span><span class="sxs-lookup"><span data-stu-id="43bc9-569">Read</span></span> | <span data-ttu-id="43bc9-570">Compose</span><span class="sxs-lookup"><span data-stu-id="43bc9-570">Compose</span></span> |

##### <a name="example"></a><span data-ttu-id="43bc9-571">例</span><span class="sxs-lookup"><span data-stu-id="43bc9-571">Example</span></span>

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

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="43bc9-572">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="43bc9-572">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="43bc9-573">ユーザーと Office アドインを識別するトークンを取得します。</span><span class="sxs-lookup"><span data-stu-id="43bc9-573">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="43bc9-574">`getUserIdentityTokenAsync` メソッドは、[アドインとユーザーをサード パーティのシステムで識別して認証](/outlook/add-ins/authentication)することのできるトークンを返します。</span><span class="sxs-lookup"><span data-stu-id="43bc9-574">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="43bc9-575">パラメーター</span><span class="sxs-lookup"><span data-stu-id="43bc9-575">Parameters</span></span>

|<span data-ttu-id="43bc9-576">名前</span><span class="sxs-lookup"><span data-stu-id="43bc9-576">Name</span></span>| <span data-ttu-id="43bc9-577">型</span><span class="sxs-lookup"><span data-stu-id="43bc9-577">Type</span></span>| <span data-ttu-id="43bc9-578">属性</span><span class="sxs-lookup"><span data-stu-id="43bc9-578">Attributes</span></span>| <span data-ttu-id="43bc9-579">説明</span><span class="sxs-lookup"><span data-stu-id="43bc9-579">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="43bc9-580">function</span><span class="sxs-lookup"><span data-stu-id="43bc9-580">function</span></span>||<span data-ttu-id="43bc9-581">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="43bc9-581">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="43bc9-582">トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="43bc9-582">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="43bc9-583">エラーが発生した場合、 `asyncResult.error` および `asyncResult.diagnostics` のプロパティで追加情報が提供される場合があります。</span><span class="sxs-lookup"><span data-stu-id="43bc9-583">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="43bc9-584">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="43bc9-584">Object</span></span>| <span data-ttu-id="43bc9-585">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="43bc9-585">&lt;optional&gt;</span></span>|<span data-ttu-id="43bc9-586">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="43bc9-586">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="43bc9-587">エラー</span><span class="sxs-lookup"><span data-stu-id="43bc9-587">Errors</span></span>

|<span data-ttu-id="43bc9-588">エラー コード</span><span class="sxs-lookup"><span data-stu-id="43bc9-588">Error code</span></span>|<span data-ttu-id="43bc9-589">説明</span><span class="sxs-lookup"><span data-stu-id="43bc9-589">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="43bc9-590">要求が失敗しました。</span><span class="sxs-lookup"><span data-stu-id="43bc9-590">The request has failed.</span></span> <span data-ttu-id="43bc9-591">HTTP エラーコードの diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="43bc9-591">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="43bc9-592">Exchange サーバーがエラーを返しました。</span><span class="sxs-lookup"><span data-stu-id="43bc9-592">The Exchange server returned an error.</span></span> <span data-ttu-id="43bc9-593">詳細については、diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="43bc9-593">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="43bc9-594">ユーザーはネットワークに接続されていません。</span><span class="sxs-lookup"><span data-stu-id="43bc9-594">The user is no longer connected to the network.</span></span> <span data-ttu-id="43bc9-595">ネットワーク接続を確認し、やり直してください。</span><span class="sxs-lookup"><span data-stu-id="43bc9-595">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="43bc9-596">要件</span><span class="sxs-lookup"><span data-stu-id="43bc9-596">Requirements</span></span>

|<span data-ttu-id="43bc9-597">必要条件</span><span class="sxs-lookup"><span data-stu-id="43bc9-597">Requirement</span></span>| <span data-ttu-id="43bc9-598">値</span><span class="sxs-lookup"><span data-stu-id="43bc9-598">Value</span></span>|
|---|---|
|[<span data-ttu-id="43bc9-599">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="43bc9-599">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43bc9-600">1.0以降</span><span class="sxs-lookup"><span data-stu-id="43bc9-600">1.0</span></span>|
|[<span data-ttu-id="43bc9-601">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="43bc9-601">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43bc9-602">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43bc9-602">ReadItem</span></span>|
|[<span data-ttu-id="43bc9-603">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="43bc9-603">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="43bc9-604">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="43bc9-604">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="43bc9-605">例</span><span class="sxs-lookup"><span data-stu-id="43bc9-605">Example</span></span>

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

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="43bc9-606">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="43bc9-606">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="43bc9-607">ユーザーのメールボックスをホストしている Exchange サーバー上の Exchange Web サービス (EWS) のサービスに対して非同期の要求を行います。</span><span class="sxs-lookup"><span data-stu-id="43bc9-607">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="43bc9-608">このメソッドは、次のシナリオではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="43bc9-608">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="43bc9-609">Outlook on iOS または Android の場合</span><span class="sxs-lookup"><span data-stu-id="43bc9-609">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="43bc9-610">アドインが Gmail のメールボックスに読み込まれる場合</span><span class="sxs-lookup"><span data-stu-id="43bc9-610">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="43bc9-611">このような場合は、アドインでは [REST API を使用](/outlook/add-ins/use-rest-api)して、代わりにユーザーのメールボックスにアクセスする必要があります。</span><span class="sxs-lookup"><span data-stu-id="43bc9-611">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="43bc9-p149">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange. See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span><span class="sxs-lookup"><span data-stu-id="43bc9-p149">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange. See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="43bc9-614">`makeEwsRequestAsync` メソッドでは、フォルダー関連アイテムを要求できません。</span><span class="sxs-lookup"><span data-stu-id="43bc9-614">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="43bc9-615">XML 要求では UTF-8 エンコードを指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="43bc9-615">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="43bc9-p150">`makeEwsRequestAsync` メソッドを使用するには、アドインに **ReadWriteMailbox** アクセス許可が必要です。**ReadWriteMailbox** アクセス許可と、`makeEwsRequestAsync` メソッドで呼び出せる EWS 操作の使い方については、「[ユーザーのメールボックスへのメール アドイン アクセスのアクセス許可を指定する](/outlook/add-ins/understanding-outlook-add-in-permissions)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="43bc9-p150">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="43bc9-618">サーバー管理者は、クライアント アクセス サーバーの EWS ディレクトリで `OAuthAuthentication` を true に設定して、`makeEwsRequestAsync` メソッドで EWS 要求を行うことができるようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="43bc9-618">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="43bc9-619">バージョンの相違点</span><span class="sxs-lookup"><span data-stu-id="43bc9-619">Version differences</span></span>

<span data-ttu-id="43bc9-620">バージョン 15.0.4535.1004 より前のバージョンの Outlook で実行しているメール アプリで `makeEwsRequestAsync` メソッドを使う場合は、エンコード値を `ISO-8859-1` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="43bc9-620">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="43bc9-p151">Outlook on the web でメール アプリを実行している場合は、エンコード値を設定する必要はありません。mailbox.diagnostics.hostName プロパティを使って、メール アプリを Outlook で実行しているのか、Outlook on the web で実行しているのかを確認できます。mailbox.diagnostics.hostVersion プロパティを使って、どのバージョンの Outlook を使って実行しているかを確認できます。</span><span class="sxs-lookup"><span data-stu-id="43bc9-p151">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="43bc9-624">パラメーター</span><span class="sxs-lookup"><span data-stu-id="43bc9-624">Parameters</span></span>

|<span data-ttu-id="43bc9-625">名前</span><span class="sxs-lookup"><span data-stu-id="43bc9-625">Name</span></span>| <span data-ttu-id="43bc9-626">型</span><span class="sxs-lookup"><span data-stu-id="43bc9-626">Type</span></span>| <span data-ttu-id="43bc9-627">属性</span><span class="sxs-lookup"><span data-stu-id="43bc9-627">Attributes</span></span>| <span data-ttu-id="43bc9-628">説明</span><span class="sxs-lookup"><span data-stu-id="43bc9-628">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="43bc9-629">String</span><span class="sxs-lookup"><span data-stu-id="43bc9-629">String</span></span>||<span data-ttu-id="43bc9-630">EWS 要求です。</span><span class="sxs-lookup"><span data-stu-id="43bc9-630">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="43bc9-631">function</span><span class="sxs-lookup"><span data-stu-id="43bc9-631">function</span></span>||<span data-ttu-id="43bc9-632">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="43bc9-632">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="43bc9-p152">The XML result of the EWS call is provided as a string in the `asyncResult.value` property. If the result exceeds 1 MB in size, an error message is returned instead.</span><span class="sxs-lookup"><span data-stu-id="43bc9-p152">The XML result of the EWS call is provided as a string in the `asyncResult.value` property. If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="43bc9-635">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="43bc9-635">Object</span></span>| <span data-ttu-id="43bc9-636">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="43bc9-636">&lt;optional&gt;</span></span>|<span data-ttu-id="43bc9-637">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="43bc9-637">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="43bc9-638">要件</span><span class="sxs-lookup"><span data-stu-id="43bc9-638">Requirements</span></span>

|<span data-ttu-id="43bc9-639">要件</span><span class="sxs-lookup"><span data-stu-id="43bc9-639">Requirement</span></span>| <span data-ttu-id="43bc9-640">値</span><span class="sxs-lookup"><span data-stu-id="43bc9-640">Value</span></span>|
|---|---|
|[<span data-ttu-id="43bc9-641">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="43bc9-641">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43bc9-642">1.0</span><span class="sxs-lookup"><span data-stu-id="43bc9-642">1.0</span></span>|
|[<span data-ttu-id="43bc9-643">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="43bc9-643">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43bc9-644">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="43bc9-644">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="43bc9-645">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="43bc9-645">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="43bc9-646">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="43bc9-646">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="43bc9-647">例</span><span class="sxs-lookup"><span data-stu-id="43bc9-647">Example</span></span>

<span data-ttu-id="43bc9-648">次の例は、`GetItem` 操作を使ってアイテムの件名を取得するため、`makeEwsRequestAsync` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="43bc9-648">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="43bc9-649">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="43bc9-649">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="43bc9-650">サポートされているイベントの種類のイベント ハンドラーを削除します。</span><span class="sxs-lookup"><span data-stu-id="43bc9-650">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="43bc9-651">現在、サポートされているイベントの種類は `Office.EventType.ItemChanged` だけです。</span><span class="sxs-lookup"><span data-stu-id="43bc9-651">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="43bc9-652">パラメーター</span><span class="sxs-lookup"><span data-stu-id="43bc9-652">Parameters</span></span>

| <span data-ttu-id="43bc9-653">名前</span><span class="sxs-lookup"><span data-stu-id="43bc9-653">Name</span></span> | <span data-ttu-id="43bc9-654">型</span><span class="sxs-lookup"><span data-stu-id="43bc9-654">Type</span></span> | <span data-ttu-id="43bc9-655">属性</span><span class="sxs-lookup"><span data-stu-id="43bc9-655">Attributes</span></span> | <span data-ttu-id="43bc9-656">説明</span><span class="sxs-lookup"><span data-stu-id="43bc9-656">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="43bc9-657">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="43bc9-657">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="43bc9-658">ハンドラーを取り消すイベント。</span><span class="sxs-lookup"><span data-stu-id="43bc9-658">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="43bc9-659">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="43bc9-659">Object</span></span> | <span data-ttu-id="43bc9-660">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="43bc9-660">&lt;optional&gt;</span></span> | <span data-ttu-id="43bc9-661">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="43bc9-661">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="43bc9-662">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="43bc9-662">Object</span></span> | <span data-ttu-id="43bc9-663">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="43bc9-663">&lt;optional&gt;</span></span> | <span data-ttu-id="43bc9-664">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="43bc9-664">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="43bc9-665">function</span><span class="sxs-lookup"><span data-stu-id="43bc9-665">function</span></span>| <span data-ttu-id="43bc9-666">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="43bc9-666">&lt;optional&gt;</span></span>|<span data-ttu-id="43bc9-667">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="43bc9-667">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="43bc9-668">Requirements</span><span class="sxs-lookup"><span data-stu-id="43bc9-668">Requirements</span></span>

|<span data-ttu-id="43bc9-669">要件</span><span class="sxs-lookup"><span data-stu-id="43bc9-669">Requirement</span></span>| <span data-ttu-id="43bc9-670">値</span><span class="sxs-lookup"><span data-stu-id="43bc9-670">Value</span></span>|
|---|---|
|[<span data-ttu-id="43bc9-671">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="43bc9-671">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43bc9-672">1.5</span><span class="sxs-lookup"><span data-stu-id="43bc9-672">1.5</span></span> |
|[<span data-ttu-id="43bc9-673">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="43bc9-673">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43bc9-674">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43bc9-674">ReadItem</span></span> |
|[<span data-ttu-id="43bc9-675">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="43bc9-675">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="43bc9-676">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="43bc9-676">Compose or Read</span></span>|
