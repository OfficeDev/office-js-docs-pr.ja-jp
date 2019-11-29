---
title: Office.context.mailbox - 要件セット 1.5
description: ''
ms.date: 11/27/2019
localization_priority: Priority
ms.openlocfilehash: eefeab2cf6fbe78451afae7e588640fe7f50dba4
ms.sourcegitcommit: 05a883a7fd89136301ce35aabc57638e9f563288
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/27/2019
ms.locfileid: "39629687"
---
# <a name="mailbox"></a><span data-ttu-id="0029d-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="0029d-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="0029d-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="0029d-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="0029d-104">Microsoft Outlook の Outlook アドイン オブジェクト モデルへのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="0029d-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="0029d-105">要件</span><span class="sxs-lookup"><span data-stu-id="0029d-105">Requirements</span></span>

|<span data-ttu-id="0029d-106">要件</span><span class="sxs-lookup"><span data-stu-id="0029d-106">Requirement</span></span>| <span data-ttu-id="0029d-107">値</span><span class="sxs-lookup"><span data-stu-id="0029d-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="0029d-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0029d-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0029d-109">1.0</span><span class="sxs-lookup"><span data-stu-id="0029d-109">1.0</span></span>|
|[<span data-ttu-id="0029d-110">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0029d-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0029d-111">制限あり</span><span class="sxs-lookup"><span data-stu-id="0029d-111">Restricted</span></span>|
|[<span data-ttu-id="0029d-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0029d-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0029d-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0029d-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="0029d-114">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="0029d-114">Members and methods</span></span>

| <span data-ttu-id="0029d-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="0029d-115">Member</span></span> | <span data-ttu-id="0029d-116">種類</span><span class="sxs-lookup"><span data-stu-id="0029d-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="0029d-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="0029d-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="0029d-118">メンバー</span><span class="sxs-lookup"><span data-stu-id="0029d-118">Member</span></span> |
| [<span data-ttu-id="0029d-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="0029d-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="0029d-120">メンバー</span><span class="sxs-lookup"><span data-stu-id="0029d-120">Member</span></span> |
| [<span data-ttu-id="0029d-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="0029d-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="0029d-122">メソッド</span><span class="sxs-lookup"><span data-stu-id="0029d-122">Method</span></span> |
| [<span data-ttu-id="0029d-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="0029d-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="0029d-124">メソッド</span><span class="sxs-lookup"><span data-stu-id="0029d-124">Method</span></span> |
| [<span data-ttu-id="0029d-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="0029d-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="0029d-126">メソッド</span><span class="sxs-lookup"><span data-stu-id="0029d-126">Method</span></span> |
| [<span data-ttu-id="0029d-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="0029d-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="0029d-128">メソッド</span><span class="sxs-lookup"><span data-stu-id="0029d-128">Method</span></span> |
| [<span data-ttu-id="0029d-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="0029d-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="0029d-130">メソッド</span><span class="sxs-lookup"><span data-stu-id="0029d-130">Method</span></span> |
| [<span data-ttu-id="0029d-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="0029d-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="0029d-132">メソッド</span><span class="sxs-lookup"><span data-stu-id="0029d-132">Method</span></span> |
| [<span data-ttu-id="0029d-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="0029d-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="0029d-134">メソッド</span><span class="sxs-lookup"><span data-stu-id="0029d-134">Method</span></span> |
| [<span data-ttu-id="0029d-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="0029d-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="0029d-136">メソッド</span><span class="sxs-lookup"><span data-stu-id="0029d-136">Method</span></span> |
| [<span data-ttu-id="0029d-137">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="0029d-137">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="0029d-138">メソッド</span><span class="sxs-lookup"><span data-stu-id="0029d-138">Method</span></span> |
| [<span data-ttu-id="0029d-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="0029d-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="0029d-140">メソッド</span><span class="sxs-lookup"><span data-stu-id="0029d-140">Method</span></span> |
| [<span data-ttu-id="0029d-141">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="0029d-141">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="0029d-142">メソッド</span><span class="sxs-lookup"><span data-stu-id="0029d-142">Method</span></span> |
| [<span data-ttu-id="0029d-143">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="0029d-143">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="0029d-144">メソッド</span><span class="sxs-lookup"><span data-stu-id="0029d-144">Method</span></span> |
| [<span data-ttu-id="0029d-145">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="0029d-145">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="0029d-146">メソッド</span><span class="sxs-lookup"><span data-stu-id="0029d-146">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="0029d-147">名前空間</span><span class="sxs-lookup"><span data-stu-id="0029d-147">Namespaces</span></span>

<span data-ttu-id="0029d-148">[diagnostics](Office.context.mailbox.diagnostics.md):Outlook アドインに診断情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="0029d-148">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="0029d-149">[item](Office.context.mailbox.item.md):Outlook アドインのメッセージや予定にアクセスするためのメソッドとプロパティを提供します。</span><span class="sxs-lookup"><span data-stu-id="0029d-149">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="0029d-150">[userProfile](Office.context.mailbox.userProfile.md):Outlook アドインのユーザーに関する情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="0029d-150">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="0029d-151">Members</span><span class="sxs-lookup"><span data-stu-id="0029d-151">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="0029d-152">ewsUrl: String</span><span class="sxs-lookup"><span data-stu-id="0029d-152">ewsUrl: String</span></span>

<span data-ttu-id="0029d-p101">このメール アカウントの Exchange Web サービス (EWS) エンドポイントの URL を取得します。読み取りモードのみです。</span><span class="sxs-lookup"><span data-stu-id="0029d-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="0029d-155">このメンバーは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="0029d-155">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="0029d-p102">
  `ewsUrl\` 値は、リモート サービスで、ユーザーのメールボックスに EWS 呼び出しを行うために使うことができます。たとえば、[選択したアイテムから添付ファイルを取得する](/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="0029d-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="0029d-158">アプリが閲覧モードで `ewsUrl` メンバーを呼び出すには、アプリのマニフェスト内に **ReadItem** アクセス許可が指定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="0029d-158">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="0029d-p103">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出してから、`ewsUrl` メンバーを使用する必要があります。アプリには、`saveAsync` メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="0029d-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="0029d-161">型</span><span class="sxs-lookup"><span data-stu-id="0029d-161">Type</span></span>

*   <span data-ttu-id="0029d-162">String</span><span class="sxs-lookup"><span data-stu-id="0029d-162">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0029d-163">要件</span><span class="sxs-lookup"><span data-stu-id="0029d-163">Requirements</span></span>

|<span data-ttu-id="0029d-164">要件</span><span class="sxs-lookup"><span data-stu-id="0029d-164">Requirement</span></span>| <span data-ttu-id="0029d-165">値</span><span class="sxs-lookup"><span data-stu-id="0029d-165">Value</span></span>|
|---|---|
|[<span data-ttu-id="0029d-166">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0029d-166">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0029d-167">1.0</span><span class="sxs-lookup"><span data-stu-id="0029d-167">1.0</span></span>|
|[<span data-ttu-id="0029d-168">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0029d-168">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0029d-169">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0029d-169">ReadItem</span></span>|
|[<span data-ttu-id="0029d-170">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0029d-170">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0029d-171">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0029d-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="resturl-string"></a><span data-ttu-id="0029d-172">restUrl: String</span><span class="sxs-lookup"><span data-stu-id="0029d-172">restUrl: String</span></span>

<span data-ttu-id="0029d-173">この電子メール アカウントの REST エンドポイントの URL を取得します。</span><span class="sxs-lookup"><span data-stu-id="0029d-173">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="0029d-174">`restUrl` 値は、ユーザーのメールボックスに [REST API](/outlook/rest/) 呼び出しを行うために使用できます。</span><span class="sxs-lookup"><span data-stu-id="0029d-174">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="0029d-175">構成されたカスタム REST URL を使用する Exchange 2016 以降のオンプレミスのインストールに接続されている Outlook クライアントは、`restUrl` に無効な値を返します。</span><span class="sxs-lookup"><span data-stu-id="0029d-175">Outlook clients connected to on-premises installations of Exchange 2016 or later with a custom REST URL configured will return an invalid value for `restUrl`.</span></span>

##### <a name="type"></a><span data-ttu-id="0029d-176">型</span><span class="sxs-lookup"><span data-stu-id="0029d-176">Type</span></span>

*   <span data-ttu-id="0029d-177">String</span><span class="sxs-lookup"><span data-stu-id="0029d-177">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0029d-178">要件</span><span class="sxs-lookup"><span data-stu-id="0029d-178">Requirements</span></span>

|<span data-ttu-id="0029d-179">要件</span><span class="sxs-lookup"><span data-stu-id="0029d-179">Requirement</span></span>| <span data-ttu-id="0029d-180">値</span><span class="sxs-lookup"><span data-stu-id="0029d-180">Value</span></span>|
|---|---|
|[<span data-ttu-id="0029d-181">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0029d-181">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0029d-182">1.5</span><span class="sxs-lookup"><span data-stu-id="0029d-182">1.5</span></span> |
|[<span data-ttu-id="0029d-183">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0029d-183">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0029d-184">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0029d-184">ReadItem</span></span>|
|[<span data-ttu-id="0029d-185">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0029d-185">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0029d-186">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0029d-186">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="0029d-187">メソッド</span><span class="sxs-lookup"><span data-stu-id="0029d-187">Methods</span></span>

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="0029d-188">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="0029d-188">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="0029d-189">サポートされているイベントのイベント ハンドラーを追加します。</span><span class="sxs-lookup"><span data-stu-id="0029d-189">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="0029d-190">現在サポートされている唯一のイベントの種類は `Office.EventType.ItemChanged` です。これはユーザーが新しいアイテムを選択したときに呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="0029d-190">Currently the only supported event type is `Office.EventType.ItemChanged`, which is invoked when the user selects a new item.</span></span> <span data-ttu-id="0029d-191">このイベントは、ピン留め可能な作業ウィンドウを実装するアドインで使用され、現在選択されているアイテムに基づいて作業ウィンドウ UI をアドインで更新できるようにします。</span><span class="sxs-lookup"><span data-stu-id="0029d-191">This event is used by add-ins that implement a pinnable task pane, and allows the add-in to refresh the task pane UI based on the currently selected item.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0029d-192">パラメーター</span><span class="sxs-lookup"><span data-stu-id="0029d-192">Parameters</span></span>

| <span data-ttu-id="0029d-193">名前</span><span class="sxs-lookup"><span data-stu-id="0029d-193">Name</span></span> | <span data-ttu-id="0029d-194">型</span><span class="sxs-lookup"><span data-stu-id="0029d-194">Type</span></span> | <span data-ttu-id="0029d-195">属性</span><span class="sxs-lookup"><span data-stu-id="0029d-195">Attributes</span></span> | <span data-ttu-id="0029d-196">説明</span><span class="sxs-lookup"><span data-stu-id="0029d-196">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="0029d-197">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="0029d-197">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="0029d-198">ハンドラーを呼び出す必要のあるイベント。</span><span class="sxs-lookup"><span data-stu-id="0029d-198">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="0029d-199">Function</span><span class="sxs-lookup"><span data-stu-id="0029d-199">Function</span></span> || <span data-ttu-id="0029d-p105">イベントを処理する関数。関数は、オブジェクト リテラルである単一パラメーターを受け入れる必要があります。パラメーターの `type` プロパティは、`addHandlerAsync` に渡される `eventType` パラメーターと一致します。</span><span class="sxs-lookup"><span data-stu-id="0029d-p105">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="0029d-203">Object</span><span class="sxs-lookup"><span data-stu-id="0029d-203">Object</span></span> | <span data-ttu-id="0029d-204">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0029d-204">&lt;optional&gt;</span></span> | <span data-ttu-id="0029d-205">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="0029d-205">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="0029d-206">Object</span><span class="sxs-lookup"><span data-stu-id="0029d-206">Object</span></span> | <span data-ttu-id="0029d-207">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0029d-207">&lt;optional&gt;</span></span> | <span data-ttu-id="0029d-208">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="0029d-208">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="0029d-209">function</span><span class="sxs-lookup"><span data-stu-id="0029d-209">function</span></span>| <span data-ttu-id="0029d-210">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0029d-210">&lt;optional&gt;</span></span>|<span data-ttu-id="0029d-211">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="0029d-211">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0029d-212">要件</span><span class="sxs-lookup"><span data-stu-id="0029d-212">Requirements</span></span>

|<span data-ttu-id="0029d-213">要件</span><span class="sxs-lookup"><span data-stu-id="0029d-213">Requirement</span></span>| <span data-ttu-id="0029d-214">値</span><span class="sxs-lookup"><span data-stu-id="0029d-214">Value</span></span>|
|---|---|
|[<span data-ttu-id="0029d-215">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0029d-215">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0029d-216">1.5</span><span class="sxs-lookup"><span data-stu-id="0029d-216">1.5</span></span> |
|[<span data-ttu-id="0029d-217">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0029d-217">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0029d-218">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0029d-218">ReadItem</span></span> |
|[<span data-ttu-id="0029d-219">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0029d-219">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0029d-220">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0029d-220">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0029d-221">例</span><span class="sxs-lookup"><span data-stu-id="0029d-221">Example</span></span>

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

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="0029d-222">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="0029d-222">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="0029d-223">REST 形式のアイテム ID を EWS 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="0029d-223">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="0029d-224">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="0029d-224">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="0029d-p106">REST API ([Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) や [Microsoft Graph](https://graph.microsoft.io/) など) で取得されたアイテム ID は、Exchange Web サービス (EWS) に使用される形式とは異なる形式を使用します。`convertToEwsId` メソッドは、REST 形式の ID を EWS 用の適切な形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="0029d-p106">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0029d-227">パラメーター</span><span class="sxs-lookup"><span data-stu-id="0029d-227">Parameters</span></span>

|<span data-ttu-id="0029d-228">名前</span><span class="sxs-lookup"><span data-stu-id="0029d-228">Name</span></span>| <span data-ttu-id="0029d-229">種類</span><span class="sxs-lookup"><span data-stu-id="0029d-229">Type</span></span>| <span data-ttu-id="0029d-230">説明</span><span class="sxs-lookup"><span data-stu-id="0029d-230">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="0029d-231">String</span><span class="sxs-lookup"><span data-stu-id="0029d-231">String</span></span>|<span data-ttu-id="0029d-232">Outlook REST API 形式のアイテム ID</span><span class="sxs-lookup"><span data-stu-id="0029d-232">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="0029d-233">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="0029d-233">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.5)|<span data-ttu-id="0029d-234">アイテム ID の取得に使用された Outlook REST API のバージョンを示す値。</span><span class="sxs-lookup"><span data-stu-id="0029d-234">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0029d-235">要件</span><span class="sxs-lookup"><span data-stu-id="0029d-235">Requirements</span></span>

|<span data-ttu-id="0029d-236">要件</span><span class="sxs-lookup"><span data-stu-id="0029d-236">Requirement</span></span>| <span data-ttu-id="0029d-237">値</span><span class="sxs-lookup"><span data-stu-id="0029d-237">Value</span></span>|
|---|---|
|[<span data-ttu-id="0029d-238">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0029d-238">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0029d-239">1.3</span><span class="sxs-lookup"><span data-stu-id="0029d-239">1.3</span></span>|
|[<span data-ttu-id="0029d-240">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0029d-240">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0029d-241">制限あり</span><span class="sxs-lookup"><span data-stu-id="0029d-241">Restricted</span></span>|
|[<span data-ttu-id="0029d-242">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0029d-242">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0029d-243">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0029d-243">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0029d-244">戻り値:</span><span class="sxs-lookup"><span data-stu-id="0029d-244">Returns:</span></span>

<span data-ttu-id="0029d-245">型:String</span><span class="sxs-lookup"><span data-stu-id="0029d-245">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="0029d-246">例</span><span class="sxs-lookup"><span data-stu-id="0029d-246">Example</span></span>

```js
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-15"></a><span data-ttu-id="0029d-247">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.5)}</span><span class="sxs-lookup"><span data-stu-id="0029d-247">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.5)}</span></span>

<span data-ttu-id="0029d-248">クライアントのローカル時間で時間情報が含まれている辞書を取得します。</span><span class="sxs-lookup"><span data-stu-id="0029d-248">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="0029d-p107">Outlook on the web または Outlook デスクトップのメールアプリでは、日付と時刻に異なるタイムゾーンを使用できます。Outlook デスクトップは、クライアント コンピューターのタイム ゾーンを使用します。Outlook on the web は、Exchange 管理センター (EAC) で設定されたタイム ゾーンを使用します。日付と時刻の値は、ユーザー インターフェイスに表示される値が、常にユーザーが期待するタイム ゾーンと一致するように処理する必要があります。</span><span class="sxs-lookup"><span data-stu-id="0029d-p107">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times. Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="0029d-p108">Outlook デスクトップ クライアントでメール アプリを実行している場合、`convertToLocalClientTime` メソッドは、クライアント コンピューターのタイム ゾーンに設定された値のディクショナリ オブジェクトを返します。Outlook on the web でメール アプリを実行している場合、`convertToLocalClientTime` メソッドは、EAC で指定したタイム ゾーンに設定された値のディクショナリ オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="0029d-p108">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0029d-254">パラメーター</span><span class="sxs-lookup"><span data-stu-id="0029d-254">Parameters</span></span>

|<span data-ttu-id="0029d-255">名前</span><span class="sxs-lookup"><span data-stu-id="0029d-255">Name</span></span>| <span data-ttu-id="0029d-256">種類</span><span class="sxs-lookup"><span data-stu-id="0029d-256">Type</span></span>| <span data-ttu-id="0029d-257">説明</span><span class="sxs-lookup"><span data-stu-id="0029d-257">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="0029d-258">Date</span><span class="sxs-lookup"><span data-stu-id="0029d-258">Date</span></span>|<span data-ttu-id="0029d-259">日付オブジェクト</span><span class="sxs-lookup"><span data-stu-id="0029d-259">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0029d-260">要件</span><span class="sxs-lookup"><span data-stu-id="0029d-260">Requirements</span></span>

|<span data-ttu-id="0029d-261">要件</span><span class="sxs-lookup"><span data-stu-id="0029d-261">Requirement</span></span>| <span data-ttu-id="0029d-262">値</span><span class="sxs-lookup"><span data-stu-id="0029d-262">Value</span></span>|
|---|---|
|[<span data-ttu-id="0029d-263">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0029d-263">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0029d-264">1.0</span><span class="sxs-lookup"><span data-stu-id="0029d-264">1.0</span></span>|
|[<span data-ttu-id="0029d-265">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0029d-265">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0029d-266">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0029d-266">ReadItem</span></span>|
|[<span data-ttu-id="0029d-267">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0029d-267">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0029d-268">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0029d-268">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0029d-269">戻り値:</span><span class="sxs-lookup"><span data-stu-id="0029d-269">Returns:</span></span>

<span data-ttu-id="0029d-270">型:[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="0029d-270">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.5)</span></span>

<br>

---
---

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="0029d-271">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="0029d-271">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="0029d-272">EWS 形式のアイテム ID を REST 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="0029d-272">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="0029d-273">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="0029d-273">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="0029d-p109">EWS または `itemId` プロパティで取得されるアイテム ID は、REST API ([Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) や [Microsoft Graph](https://graph.microsoft.io/) など) に使用される形式とは異なる形式を使用します。`convertToRestId` メソッドは、EWS 形式の ID を REST 用の適切な形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="0029d-p109">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0029d-276">パラメーター</span><span class="sxs-lookup"><span data-stu-id="0029d-276">Parameters</span></span>

|<span data-ttu-id="0029d-277">名前</span><span class="sxs-lookup"><span data-stu-id="0029d-277">Name</span></span>| <span data-ttu-id="0029d-278">種類</span><span class="sxs-lookup"><span data-stu-id="0029d-278">Type</span></span>| <span data-ttu-id="0029d-279">説明</span><span class="sxs-lookup"><span data-stu-id="0029d-279">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="0029d-280">String</span><span class="sxs-lookup"><span data-stu-id="0029d-280">String</span></span>|<span data-ttu-id="0029d-281">Exchange Web サービス (EWS) 形式のアイテム ID</span><span class="sxs-lookup"><span data-stu-id="0029d-281">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="0029d-282">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="0029d-282">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.5)|<span data-ttu-id="0029d-283">変換後の ID を使用する Outlook REST API のバージョンを示す値。</span><span class="sxs-lookup"><span data-stu-id="0029d-283">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0029d-284">要件</span><span class="sxs-lookup"><span data-stu-id="0029d-284">Requirements</span></span>

|<span data-ttu-id="0029d-285">要件</span><span class="sxs-lookup"><span data-stu-id="0029d-285">Requirement</span></span>| <span data-ttu-id="0029d-286">値</span><span class="sxs-lookup"><span data-stu-id="0029d-286">Value</span></span>|
|---|---|
|[<span data-ttu-id="0029d-287">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0029d-287">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0029d-288">1.3</span><span class="sxs-lookup"><span data-stu-id="0029d-288">1.3</span></span>|
|[<span data-ttu-id="0029d-289">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0029d-289">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0029d-290">制限あり</span><span class="sxs-lookup"><span data-stu-id="0029d-290">Restricted</span></span>|
|[<span data-ttu-id="0029d-291">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0029d-291">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0029d-292">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0029d-292">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0029d-293">戻り値:</span><span class="sxs-lookup"><span data-stu-id="0029d-293">Returns:</span></span>

<span data-ttu-id="0029d-294">型:String</span><span class="sxs-lookup"><span data-stu-id="0029d-294">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="0029d-295">例</span><span class="sxs-lookup"><span data-stu-id="0029d-295">Example</span></span>

```js
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="0029d-296">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="0029d-296">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="0029d-297">時間情報が含まれているディクショナリから日付オブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="0029d-297">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="0029d-298">`convertToUtcClientTime` メソッドは、ローカルの日付と時刻を含むディクショナリを、ローカルの日付と時刻の正しい値を持つ日付オブジェクトに変換します。</span><span class="sxs-lookup"><span data-stu-id="0029d-298">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0029d-299">パラメーター</span><span class="sxs-lookup"><span data-stu-id="0029d-299">Parameters</span></span>

|<span data-ttu-id="0029d-300">名前</span><span class="sxs-lookup"><span data-stu-id="0029d-300">Name</span></span>| <span data-ttu-id="0029d-301">種類</span><span class="sxs-lookup"><span data-stu-id="0029d-301">Type</span></span>| <span data-ttu-id="0029d-302">説明</span><span class="sxs-lookup"><span data-stu-id="0029d-302">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="0029d-303">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="0029d-303">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.5)|<span data-ttu-id="0029d-304">変換するローカル時刻の値。</span><span class="sxs-lookup"><span data-stu-id="0029d-304">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0029d-305">要件</span><span class="sxs-lookup"><span data-stu-id="0029d-305">Requirements</span></span>

|<span data-ttu-id="0029d-306">要件</span><span class="sxs-lookup"><span data-stu-id="0029d-306">Requirement</span></span>| <span data-ttu-id="0029d-307">値</span><span class="sxs-lookup"><span data-stu-id="0029d-307">Value</span></span>|
|---|---|
|[<span data-ttu-id="0029d-308">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0029d-308">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0029d-309">1.0</span><span class="sxs-lookup"><span data-stu-id="0029d-309">1.0</span></span>|
|[<span data-ttu-id="0029d-310">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0029d-310">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0029d-311">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0029d-311">ReadItem</span></span>|
|[<span data-ttu-id="0029d-312">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0029d-312">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0029d-313">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0029d-313">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0029d-314">戻り値:</span><span class="sxs-lookup"><span data-stu-id="0029d-314">Returns:</span></span>

<span data-ttu-id="0029d-315">時間が UTC で表現された日付オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="0029d-315">A Date object with the time expressed in UTC.</span></span>

<span data-ttu-id="0029d-316">型: Date</span><span class="sxs-lookup"><span data-stu-id="0029d-316">Type: Date</span></span>

##### <a name="example"></a><span data-ttu-id="0029d-317">例</span><span class="sxs-lookup"><span data-stu-id="0029d-317">Example</span></span>

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

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="0029d-318">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="0029d-318">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="0029d-319">既存の予定を表示します。</span><span class="sxs-lookup"><span data-stu-id="0029d-319">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="0029d-320">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="0029d-320">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="0029d-321">`displayAppointmentForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存の予定を開きます。</span><span class="sxs-lookup"><span data-stu-id="0029d-321">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="0029d-p110">Outlook on Mac では、このメソッドを使用して定期的な系列に含まれない単発の予定や定期的な系列のマスター予定を表示できます。ただし、系列のインスタンスは表示できません。これは、Outlook on Mac では定期的な系列のインスタンスのプロパティ (アイテム ID を含む) にアクセスできないためです。</span><span class="sxs-lookup"><span data-stu-id="0029d-p110">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="0029d-324">Outlook on the web では、このメソッドはフォームの本文が 32KB 以下の文字数の場合にのみ指定のフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="0029d-324">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="0029d-325">指定のアイテム識別子が既存の予定を表していない場合は、クライアント コンピューターまたはデバイスで空のウィンドウが開き、エラー メッセージは返されません。</span><span class="sxs-lookup"><span data-stu-id="0029d-325">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0029d-326">パラメーター</span><span class="sxs-lookup"><span data-stu-id="0029d-326">Parameters</span></span>

|<span data-ttu-id="0029d-327">名前</span><span class="sxs-lookup"><span data-stu-id="0029d-327">Name</span></span>| <span data-ttu-id="0029d-328">種類</span><span class="sxs-lookup"><span data-stu-id="0029d-328">Type</span></span>| <span data-ttu-id="0029d-329">説明</span><span class="sxs-lookup"><span data-stu-id="0029d-329">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="0029d-330">String</span><span class="sxs-lookup"><span data-stu-id="0029d-330">String</span></span>|<span data-ttu-id="0029d-331">既存の予定の Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="0029d-331">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0029d-332">要件</span><span class="sxs-lookup"><span data-stu-id="0029d-332">Requirements</span></span>

|<span data-ttu-id="0029d-333">要件</span><span class="sxs-lookup"><span data-stu-id="0029d-333">Requirement</span></span>| <span data-ttu-id="0029d-334">値</span><span class="sxs-lookup"><span data-stu-id="0029d-334">Value</span></span>|
|---|---|
|[<span data-ttu-id="0029d-335">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0029d-335">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0029d-336">1.0</span><span class="sxs-lookup"><span data-stu-id="0029d-336">1.0</span></span>|
|[<span data-ttu-id="0029d-337">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0029d-337">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0029d-338">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0029d-338">ReadItem</span></span>|
|[<span data-ttu-id="0029d-339">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0029d-339">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0029d-340">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0029d-340">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0029d-341">例</span><span class="sxs-lookup"><span data-stu-id="0029d-341">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

<br>

---
---

#### <a name="displaymessageformitemid"></a><span data-ttu-id="0029d-342">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="0029d-342">displayMessageForm(itemId)</span></span>

<span data-ttu-id="0029d-343">既存のメッセージを表示します。</span><span class="sxs-lookup"><span data-stu-id="0029d-343">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="0029d-344">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="0029d-344">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="0029d-345">`displayMessageForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存のメッセージを開きます。</span><span class="sxs-lookup"><span data-stu-id="0029d-345">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="0029d-346">Outlook on the web では、このメソッドはフォームの本文が 32 KB 以下の文字数の場合にのみ指定のフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="0029d-346">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="0029d-347">指定のアイテム識別子が既存のメッセージを表していない場合は、クライアント コンピューターにはメッセージは表示されず、エラー メッセージも返されません。</span><span class="sxs-lookup"><span data-stu-id="0029d-347">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="0029d-p111">予定を表す `itemId` を含む `displayMessageForm` を使用しないでください。既存の予定を表示するには、`displayAppointmentForm` メソッドを使用します。新しい予定を作成するフォームを表示するには、`displayNewAppointmentForm` メソッドを使用します。</span><span class="sxs-lookup"><span data-stu-id="0029d-p111">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0029d-350">パラメーター</span><span class="sxs-lookup"><span data-stu-id="0029d-350">Parameters</span></span>

|<span data-ttu-id="0029d-351">名前</span><span class="sxs-lookup"><span data-stu-id="0029d-351">Name</span></span>| <span data-ttu-id="0029d-352">種類</span><span class="sxs-lookup"><span data-stu-id="0029d-352">Type</span></span>| <span data-ttu-id="0029d-353">説明</span><span class="sxs-lookup"><span data-stu-id="0029d-353">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="0029d-354">String</span><span class="sxs-lookup"><span data-stu-id="0029d-354">String</span></span>|<span data-ttu-id="0029d-355">既存のメッセージの Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="0029d-355">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0029d-356">要件</span><span class="sxs-lookup"><span data-stu-id="0029d-356">Requirements</span></span>

|<span data-ttu-id="0029d-357">要件</span><span class="sxs-lookup"><span data-stu-id="0029d-357">Requirement</span></span>| <span data-ttu-id="0029d-358">値</span><span class="sxs-lookup"><span data-stu-id="0029d-358">Value</span></span>|
|---|---|
|[<span data-ttu-id="0029d-359">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0029d-359">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0029d-360">1.0</span><span class="sxs-lookup"><span data-stu-id="0029d-360">1.0</span></span>|
|[<span data-ttu-id="0029d-361">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0029d-361">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0029d-362">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0029d-362">ReadItem</span></span>|
|[<span data-ttu-id="0029d-363">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0029d-363">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0029d-364">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0029d-364">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0029d-365">例</span><span class="sxs-lookup"><span data-stu-id="0029d-365">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

<br>

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="0029d-366">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="0029d-366">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="0029d-367">新しい予定を作成するためのフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="0029d-367">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="0029d-368">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="0029d-368">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="0029d-p112">`displayNewAppointmentForm` メソッドを使用すると、ユーザーが新しい予定または会議を作成できるフォームが開きます。パラメーターを指定すると、予定のフォーム フィールドにパラメーターの内容が自動的に設定されます。</span><span class="sxs-lookup"><span data-stu-id="0029d-p112">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="0029d-p113">Outlook on the web およびモバイル デバイスでは、このメソッドは常に出席者フィールドが含まれるフォームを表示します。入力引数として出席者を指定しないと、このメソッドは **[保存]** ボタンのあるフォームを表示します。出席者を指定した場合には、フォームにその出席者と **[送信]** ボタンが表示されます。</span><span class="sxs-lookup"><span data-stu-id="0029d-p113">In Outlook on the web and mobile devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="0029d-p114">Outlook リッチ クライアントと Outlook RT で、`requiredAttendees`、`optionalAttendees`、または `resources` パラメーターに出席者またはリソースを指定し、このメソッドを実行すると、**[送信]** ボタンがある会議フォームが表示されます。受信者を指定せずにこのメソッドを実行すると、**[保存して閉じる]** ボタンがある予定フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="0029d-p114">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="0029d-376">パラメーターのいずれかが指定のサイズ制限を超える場合、または不明なパラメーター名が指定されている場合は、例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="0029d-376">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0029d-377">パラメーター</span><span class="sxs-lookup"><span data-stu-id="0029d-377">Parameters</span></span>

|<span data-ttu-id="0029d-378">名前</span><span class="sxs-lookup"><span data-stu-id="0029d-378">Name</span></span>| <span data-ttu-id="0029d-379">種類</span><span class="sxs-lookup"><span data-stu-id="0029d-379">Type</span></span>| <span data-ttu-id="0029d-380">説明</span><span class="sxs-lookup"><span data-stu-id="0029d-380">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="0029d-381">Object</span><span class="sxs-lookup"><span data-stu-id="0029d-381">Object</span></span> | <span data-ttu-id="0029d-382">新しい予定を記述するパラメーターのディクショナリ。</span><span class="sxs-lookup"><span data-stu-id="0029d-382">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="0029d-383">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)&gt;</span><span class="sxs-lookup"><span data-stu-id="0029d-383">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)&gt;</span></span> | <span data-ttu-id="0029d-p115">予定に必要な各出席者について、メール アドレスを含む文字列の配列、または `EmailAddressDetails` オブジェクトを含む配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="0029d-p115">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="0029d-386">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)&gt;</span><span class="sxs-lookup"><span data-stu-id="0029d-386">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)&gt;</span></span> | <span data-ttu-id="0029d-p116">予定の各任意出席者について、メール アドレスを含む文字列の配列、または `EmailAddressDetails` オブジェクトを含む配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="0029d-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="0029d-389">日付</span><span class="sxs-lookup"><span data-stu-id="0029d-389">Date</span></span> | <span data-ttu-id="0029d-390">予定の開始日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="0029d-390">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="0029d-391">Date</span><span class="sxs-lookup"><span data-stu-id="0029d-391">Date</span></span> | <span data-ttu-id="0029d-392">予定の終了日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="0029d-392">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="0029d-393">String</span><span class="sxs-lookup"><span data-stu-id="0029d-393">String</span></span> | <span data-ttu-id="0029d-p117">予定の場所を含む文字列。文字列は最大 255 文字に制限されます。</span><span class="sxs-lookup"><span data-stu-id="0029d-p117">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="0029d-396">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="0029d-396">Array.&lt;String&gt;</span></span> | <span data-ttu-id="0029d-p118">予定に必要なリソースを含む文字列の配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="0029d-p118">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="0029d-399">String</span><span class="sxs-lookup"><span data-stu-id="0029d-399">String</span></span> | <span data-ttu-id="0029d-p119">予定の件名を含む文字列です。文字列は最大 255 文字に制限されます。</span><span class="sxs-lookup"><span data-stu-id="0029d-p119">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="0029d-402">String</span><span class="sxs-lookup"><span data-stu-id="0029d-402">String</span></span> | <span data-ttu-id="0029d-p120">予定の本文。本文の内容は、最大サイズが 32 KB に制限されます。</span><span class="sxs-lookup"><span data-stu-id="0029d-p120">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="0029d-405">要件</span><span class="sxs-lookup"><span data-stu-id="0029d-405">Requirements</span></span>

|<span data-ttu-id="0029d-406">要件</span><span class="sxs-lookup"><span data-stu-id="0029d-406">Requirement</span></span>| <span data-ttu-id="0029d-407">値</span><span class="sxs-lookup"><span data-stu-id="0029d-407">Value</span></span>|
|---|---|
|[<span data-ttu-id="0029d-408">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0029d-408">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0029d-409">1.0</span><span class="sxs-lookup"><span data-stu-id="0029d-409">1.0</span></span>|
|[<span data-ttu-id="0029d-410">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0029d-410">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0029d-411">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0029d-411">ReadItem</span></span>|
|[<span data-ttu-id="0029d-412">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0029d-412">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0029d-413">読み取り</span><span class="sxs-lookup"><span data-stu-id="0029d-413">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0029d-414">例</span><span class="sxs-lookup"><span data-stu-id="0029d-414">Example</span></span>

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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="0029d-415">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="0029d-415">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="0029d-416">REST API または Exchange Web サービスを呼び出すために使用するトークンを含む文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="0029d-416">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="0029d-p121">`getCallbackTokenAsync` メソッドは、ユーザーのメールボックスをホストする Exchange Server から不透明なトークンを取得する非同期の呼び出しを行います。コールバック トークンの有効期間は 5 分です。</span><span class="sxs-lookup"><span data-stu-id="0029d-p121">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="0029d-419">可能であれば、アドインでは Exchange Web サービスの代わりに REST API を使用することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="0029d-419">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span>

<span data-ttu-id="0029d-420">閲覧モードで `getCallbackTokenAsync` メソッドを呼び出すには、**ReadItem** の最小限のアクセス許可レベルが必要です。</span><span class="sxs-lookup"><span data-stu-id="0029d-420">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="0029d-421">作成モードで `getCallbackTokenAsync` を呼び出すには、アイテムを保存しておく必要があります。</span><span class="sxs-lookup"><span data-stu-id="0029d-421">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="0029d-422">[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドは、**ReadWriteItem** の最小限のアクセス許可レベルが必要です。</span><span class="sxs-lookup"><span data-stu-id="0029d-422">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

<span data-ttu-id="0029d-423">**REST トークン**</span><span class="sxs-lookup"><span data-stu-id="0029d-423">**REST Tokens**</span></span>

<span data-ttu-id="0029d-p123">REST トークンが要求された場合 (`options.isRest = true`)、結果トークンは Exchange Web サービスの呼び出しを認証するためには機能しません。アドインがマニフェストで [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) アクセス許可を指定していない限り、トークンの範囲は現在のアイテムとその添付ファイルへの読み取り専用アクセスに制限されます。`ReadWriteMailbox` アクセス許可が指定されている場合は、結果トークンは、メールを送信する機能など、メール、カレンダー、連絡先への読み取り/書き込みアクセスを付与します。</span><span class="sxs-lookup"><span data-stu-id="0029d-p123">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="0029d-427">アドインでは、`restUrl` プロパティを使用して、REST API 呼び出しを行うときに使用する正しい URL を決定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="0029d-427">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="0029d-428">**EWS トークン**</span><span class="sxs-lookup"><span data-stu-id="0029d-428">**EWS Tokens**</span></span>

<span data-ttu-id="0029d-p124">EWS トークンが要求された場合 (`options.isRest = false`)、結果トークンは REST API 呼び出しを認証するためには機能しません。トークンの範囲は、現在のアイテムへのアクセスに制限されます。</span><span class="sxs-lookup"><span data-stu-id="0029d-p124">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="0029d-431">アドインでは、`ewsUrl` プロパティを使用して、EWS 呼び出しを行うときに使用する正しい URL を決定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="0029d-431">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

<span data-ttu-id="0029d-432">トークンと、添付ファイル識別子またはアイテム識別子の両方をサードパーティ システムに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="0029d-432">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="0029d-433">サードパーティ システムは、トークンをベアラー承認トークンとして使用し、Exchange Web サービス (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) 操作または [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) 操作を呼び出して、添付ファイルまたはアイテムを返します。</span><span class="sxs-lookup"><span data-stu-id="0029d-433">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item.</span></span> <span data-ttu-id="0029d-434">たとえば、[選択したアイテムから添付ファイルを取得する](/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="0029d-434">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

##### <a name="parameters"></a><span data-ttu-id="0029d-435">パラメーター</span><span class="sxs-lookup"><span data-stu-id="0029d-435">Parameters</span></span>

|<span data-ttu-id="0029d-436">名前</span><span class="sxs-lookup"><span data-stu-id="0029d-436">Name</span></span>| <span data-ttu-id="0029d-437">型</span><span class="sxs-lookup"><span data-stu-id="0029d-437">Type</span></span>| <span data-ttu-id="0029d-438">属性</span><span class="sxs-lookup"><span data-stu-id="0029d-438">Attributes</span></span>| <span data-ttu-id="0029d-439">説明</span><span class="sxs-lookup"><span data-stu-id="0029d-439">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="0029d-440">Object</span><span class="sxs-lookup"><span data-stu-id="0029d-440">Object</span></span> | <span data-ttu-id="0029d-441">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0029d-441">&lt;optional&gt;</span></span> | <span data-ttu-id="0029d-442">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="0029d-442">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="0029d-443">Boolean</span><span class="sxs-lookup"><span data-stu-id="0029d-443">Boolean</span></span> |  <span data-ttu-id="0029d-444">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="0029d-444">&lt;optional&gt;</span></span> | <span data-ttu-id="0029d-p126">提供されたトークンを Outlook REST API または Exchange Web サービスに使用するかどうかを決定します。既定値は、`false` です。</span><span class="sxs-lookup"><span data-stu-id="0029d-p126">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="0029d-447">Object</span><span class="sxs-lookup"><span data-stu-id="0029d-447">Object</span></span> |  <span data-ttu-id="0029d-448">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="0029d-448">&lt;optional&gt;</span></span> | <span data-ttu-id="0029d-449">非同期メソッドに渡される状態データ。</span><span class="sxs-lookup"><span data-stu-id="0029d-449">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="0029d-450">function</span><span class="sxs-lookup"><span data-stu-id="0029d-450">function</span></span>||<span data-ttu-id="0029d-451">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="0029d-451">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="0029d-452">トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="0029d-452">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="0029d-453">エラーが発生した場合、 `asyncResult.error` および `asyncResult.diagnostics` のプロパティで追加情報が提供される場合があります。</span><span class="sxs-lookup"><span data-stu-id="0029d-453">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|

##### <a name="errors"></a><span data-ttu-id="0029d-454">エラー</span><span class="sxs-lookup"><span data-stu-id="0029d-454">Errors</span></span>

|<span data-ttu-id="0029d-455">エラー コード</span><span class="sxs-lookup"><span data-stu-id="0029d-455">Error code</span></span>|<span data-ttu-id="0029d-456">説明</span><span class="sxs-lookup"><span data-stu-id="0029d-456">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="0029d-457">要求が失敗しました。</span><span class="sxs-lookup"><span data-stu-id="0029d-457">The request has failed.</span></span> <span data-ttu-id="0029d-458">HTTP エラーコードの diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="0029d-458">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="0029d-459">Exchange サーバーがエラーを返しました。</span><span class="sxs-lookup"><span data-stu-id="0029d-459">The Exchange server returned an error.</span></span> <span data-ttu-id="0029d-460">詳細については、diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="0029d-460">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="0029d-461">ユーザーはネットワークに接続されていません。</span><span class="sxs-lookup"><span data-stu-id="0029d-461">The user is no longer connected to the network.</span></span> <span data-ttu-id="0029d-462">ネットワーク接続を確認し、やり直してください。</span><span class="sxs-lookup"><span data-stu-id="0029d-462">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0029d-463">要件</span><span class="sxs-lookup"><span data-stu-id="0029d-463">Requirements</span></span>

|<span data-ttu-id="0029d-464">要件</span><span class="sxs-lookup"><span data-stu-id="0029d-464">Requirement</span></span>| <span data-ttu-id="0029d-465">値</span><span class="sxs-lookup"><span data-stu-id="0029d-465">Value</span></span>|
|---|---|
|[<span data-ttu-id="0029d-466">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0029d-466">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0029d-467">1.5</span><span class="sxs-lookup"><span data-stu-id="0029d-467">1.5</span></span> |
|[<span data-ttu-id="0029d-468">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0029d-468">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0029d-469">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0029d-469">ReadItem</span></span>|
|[<span data-ttu-id="0029d-470">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0029d-470">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0029d-471">新規作成と閲覧</span><span class="sxs-lookup"><span data-stu-id="0029d-471">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="0029d-472">例</span><span class="sxs-lookup"><span data-stu-id="0029d-472">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="0029d-473">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="0029d-473">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="0029d-474">Exchange Server から添付ファイルやアイテムを取得するために使うトークンを含む文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="0029d-474">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="0029d-p130">`getCallbackTokenAsync` メソッドは、ユーザーのメールボックスをホストする Exchange Server から不透明なトークンを取得する非同期の呼び出しを行います。コールバック トークンの有効期間は 5 分です。</span><span class="sxs-lookup"><span data-stu-id="0029d-p130">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="0029d-477">トークンと、添付ファイル識別子またはアイテム識別子の両方をサードパーティ システムに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="0029d-477">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="0029d-478">サードパーティ システムは、トークンをベアラー承認トークンとして使用し、Exchange Web サービス (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) 操作または [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) 操作を呼び出して、添付ファイルまたはアイテムを返します。</span><span class="sxs-lookup"><span data-stu-id="0029d-478">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item.</span></span> <span data-ttu-id="0029d-479">たとえば、[選択したアイテムから添付ファイルを取得する](/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="0029d-479">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="0029d-480">閲覧モードで `getCallbackTokenAsync` メソッドを呼び出すには、**ReadItem** の最小限のアクセス許可レベルが必要です。</span><span class="sxs-lookup"><span data-stu-id="0029d-480">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="0029d-481">作成モードで `getCallbackTokenAsync` を呼び出すには、アイテムを保存しておく必要があります。</span><span class="sxs-lookup"><span data-stu-id="0029d-481">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="0029d-482">[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドは、**ReadWriteItem** の最小限のアクセス許可レベルが必要です。</span><span class="sxs-lookup"><span data-stu-id="0029d-482">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0029d-483">パラメーター</span><span class="sxs-lookup"><span data-stu-id="0029d-483">Parameters</span></span>

|<span data-ttu-id="0029d-484">名前</span><span class="sxs-lookup"><span data-stu-id="0029d-484">Name</span></span>| <span data-ttu-id="0029d-485">型</span><span class="sxs-lookup"><span data-stu-id="0029d-485">Type</span></span>| <span data-ttu-id="0029d-486">属性</span><span class="sxs-lookup"><span data-stu-id="0029d-486">Attributes</span></span>| <span data-ttu-id="0029d-487">説明</span><span class="sxs-lookup"><span data-stu-id="0029d-487">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="0029d-488">function</span><span class="sxs-lookup"><span data-stu-id="0029d-488">function</span></span>||<span data-ttu-id="0029d-489">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="0029d-489">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="0029d-490">トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="0029d-490">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="0029d-491">エラーが発生した場合、 `asyncResult.error` および `asyncResult.diagnostics` のプロパティで追加情報が提供される場合があります。</span><span class="sxs-lookup"><span data-stu-id="0029d-491">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="0029d-492">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="0029d-492">Object</span></span>| <span data-ttu-id="0029d-493">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="0029d-493">&lt;optional&gt;</span></span>|<span data-ttu-id="0029d-494">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="0029d-494">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="0029d-495">エラー</span><span class="sxs-lookup"><span data-stu-id="0029d-495">Errors</span></span>

|<span data-ttu-id="0029d-496">エラー コード</span><span class="sxs-lookup"><span data-stu-id="0029d-496">Error code</span></span>|<span data-ttu-id="0029d-497">説明</span><span class="sxs-lookup"><span data-stu-id="0029d-497">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="0029d-498">要求が失敗しました。</span><span class="sxs-lookup"><span data-stu-id="0029d-498">The request has failed.</span></span> <span data-ttu-id="0029d-499">HTTP エラーコードの diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="0029d-499">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="0029d-500">Exchange サーバーがエラーを返しました。</span><span class="sxs-lookup"><span data-stu-id="0029d-500">The Exchange server returned an error.</span></span> <span data-ttu-id="0029d-501">詳細については、diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="0029d-501">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="0029d-502">ユーザーはネットワークに接続されていません。</span><span class="sxs-lookup"><span data-stu-id="0029d-502">The user is no longer connected to the network.</span></span> <span data-ttu-id="0029d-503">ネットワーク接続を確認し、やり直してください。</span><span class="sxs-lookup"><span data-stu-id="0029d-503">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0029d-504">要件</span><span class="sxs-lookup"><span data-stu-id="0029d-504">Requirements</span></span>

|<span data-ttu-id="0029d-505">要件</span><span class="sxs-lookup"><span data-stu-id="0029d-505">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="0029d-506">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0029d-506">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0029d-507">1.0</span><span class="sxs-lookup"><span data-stu-id="0029d-507">1.0</span></span> | <span data-ttu-id="0029d-508">1.3</span><span class="sxs-lookup"><span data-stu-id="0029d-508">1.3</span></span> |
|[<span data-ttu-id="0029d-509">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0029d-509">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0029d-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0029d-510">ReadItem</span></span> | <span data-ttu-id="0029d-511">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0029d-511">ReadItem</span></span> |
|[<span data-ttu-id="0029d-512">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0029d-512">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0029d-513">Read</span><span class="sxs-lookup"><span data-stu-id="0029d-513">Read</span></span> | <span data-ttu-id="0029d-514">Compose</span><span class="sxs-lookup"><span data-stu-id="0029d-514">Compose</span></span> |

##### <a name="example"></a><span data-ttu-id="0029d-515">例</span><span class="sxs-lookup"><span data-stu-id="0029d-515">Example</span></span>

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

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="0029d-516">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="0029d-516">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="0029d-517">ユーザーと Office アドインを識別するトークンを取得します。</span><span class="sxs-lookup"><span data-stu-id="0029d-517">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="0029d-518">`getUserIdentityTokenAsync` メソッドは、[アドインとユーザーをサード パーティのシステムで識別して認証](/outlook/add-ins/authentication)することのできるトークンを返します。</span><span class="sxs-lookup"><span data-stu-id="0029d-518">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="0029d-519">パラメーター</span><span class="sxs-lookup"><span data-stu-id="0029d-519">Parameters</span></span>

|<span data-ttu-id="0029d-520">名前</span><span class="sxs-lookup"><span data-stu-id="0029d-520">Name</span></span>| <span data-ttu-id="0029d-521">型</span><span class="sxs-lookup"><span data-stu-id="0029d-521">Type</span></span>| <span data-ttu-id="0029d-522">属性</span><span class="sxs-lookup"><span data-stu-id="0029d-522">Attributes</span></span>| <span data-ttu-id="0029d-523">説明</span><span class="sxs-lookup"><span data-stu-id="0029d-523">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="0029d-524">function</span><span class="sxs-lookup"><span data-stu-id="0029d-524">function</span></span>||<span data-ttu-id="0029d-525">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="0029d-525">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="0029d-526">トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="0029d-526">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="0029d-527">エラーが発生した場合、 `asyncResult.error` および `asyncResult.diagnostics` のプロパティで追加情報が提供される場合があります。</span><span class="sxs-lookup"><span data-stu-id="0029d-527">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="0029d-528">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="0029d-528">Object</span></span>| <span data-ttu-id="0029d-529">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="0029d-529">&lt;optional&gt;</span></span>|<span data-ttu-id="0029d-530">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="0029d-530">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="0029d-531">エラー</span><span class="sxs-lookup"><span data-stu-id="0029d-531">Errors</span></span>

|<span data-ttu-id="0029d-532">エラー コード</span><span class="sxs-lookup"><span data-stu-id="0029d-532">Error code</span></span>|<span data-ttu-id="0029d-533">説明</span><span class="sxs-lookup"><span data-stu-id="0029d-533">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="0029d-534">要求が失敗しました。</span><span class="sxs-lookup"><span data-stu-id="0029d-534">The request has failed.</span></span> <span data-ttu-id="0029d-535">HTTP エラーコードの diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="0029d-535">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="0029d-536">Exchange サーバーがエラーを返しました。</span><span class="sxs-lookup"><span data-stu-id="0029d-536">The Exchange server returned an error.</span></span> <span data-ttu-id="0029d-537">詳細については、diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="0029d-537">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="0029d-538">ユーザーはネットワークに接続されていません。</span><span class="sxs-lookup"><span data-stu-id="0029d-538">The user is no longer connected to the network.</span></span> <span data-ttu-id="0029d-539">ネットワーク接続を確認し、やり直してください。</span><span class="sxs-lookup"><span data-stu-id="0029d-539">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0029d-540">要件</span><span class="sxs-lookup"><span data-stu-id="0029d-540">Requirements</span></span>

|<span data-ttu-id="0029d-541">要件</span><span class="sxs-lookup"><span data-stu-id="0029d-541">Requirement</span></span>| <span data-ttu-id="0029d-542">値</span><span class="sxs-lookup"><span data-stu-id="0029d-542">Value</span></span>|
|---|---|
|[<span data-ttu-id="0029d-543">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0029d-543">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0029d-544">1.0</span><span class="sxs-lookup"><span data-stu-id="0029d-544">1.0</span></span>|
|[<span data-ttu-id="0029d-545">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0029d-545">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0029d-546">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0029d-546">ReadItem</span></span>|
|[<span data-ttu-id="0029d-547">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0029d-547">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0029d-548">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0029d-548">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0029d-549">例</span><span class="sxs-lookup"><span data-stu-id="0029d-549">Example</span></span>

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

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="0029d-550">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="0029d-550">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="0029d-551">ユーザーのメールボックスをホストしている Exchange サーバー上の Exchange Web サービス (EWS) のサービスに対して非同期の要求を行います。</span><span class="sxs-lookup"><span data-stu-id="0029d-551">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="0029d-552">このメソッドは、次のシナリオではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="0029d-552">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="0029d-553">Outlook on iOS または Android の場合</span><span class="sxs-lookup"><span data-stu-id="0029d-553">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="0029d-554">アドインが Gmail のメールボックスに読み込まれる場合</span><span class="sxs-lookup"><span data-stu-id="0029d-554">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="0029d-555">このような場合は、アドインでは [REST API を使用](/outlook/add-ins/use-rest-api)して、代わりにユーザーのメールボックスにアクセスする必要があります。</span><span class="sxs-lookup"><span data-stu-id="0029d-555">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="0029d-p139">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange. See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span><span class="sxs-lookup"><span data-stu-id="0029d-p139">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange. See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="0029d-558">`makeEwsRequestAsync` メソッドでは、フォルダー関連アイテムを要求できません。</span><span class="sxs-lookup"><span data-stu-id="0029d-558">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="0029d-559">XML 要求では UTF-8 エンコードを指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="0029d-559">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="0029d-p140">`makeEwsRequestAsync` メソッドを使用するには、アドインに **ReadWriteMailbox** アクセス許可が必要です。**ReadWriteMailbox** アクセス許可と、`makeEwsRequestAsync` メソッドで呼び出せる EWS 操作の使い方については、「[ユーザーのメールボックスへのメール アドイン アクセスのアクセス許可を指定する](/outlook/add-ins/understanding-outlook-add-in-permissions)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="0029d-p140">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="0029d-562">サーバー管理者は、クライアント アクセス サーバーの EWS ディレクトリで `OAuthAuthentication` を true に設定して、`makeEwsRequestAsync` メソッドで EWS 要求を行うことができるようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="0029d-562">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="0029d-563">バージョンの相違点</span><span class="sxs-lookup"><span data-stu-id="0029d-563">Version differences</span></span>

<span data-ttu-id="0029d-564">バージョン 15.0.4535.1004 より前のバージョンの Outlook で実行しているメール アプリで `makeEwsRequestAsync` メソッドを使う場合は、エンコード値を `ISO-8859-1` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="0029d-564">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="0029d-p141">Outlook on the web でメール アプリを実行している場合は、エンコード値を設定する必要はありません。mailbox.diagnostics.hostName プロパティを使って、メール アプリを Outlook で実行しているのか、Outlook on the web で実行しているのかを確認できます。mailbox.diagnostics.hostVersion プロパティを使って、どのバージョンの Outlook を使って実行しているかを確認できます。</span><span class="sxs-lookup"><span data-stu-id="0029d-p141">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0029d-568">パラメーター</span><span class="sxs-lookup"><span data-stu-id="0029d-568">Parameters</span></span>

|<span data-ttu-id="0029d-569">名前</span><span class="sxs-lookup"><span data-stu-id="0029d-569">Name</span></span>| <span data-ttu-id="0029d-570">型</span><span class="sxs-lookup"><span data-stu-id="0029d-570">Type</span></span>| <span data-ttu-id="0029d-571">属性</span><span class="sxs-lookup"><span data-stu-id="0029d-571">Attributes</span></span>| <span data-ttu-id="0029d-572">説明</span><span class="sxs-lookup"><span data-stu-id="0029d-572">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="0029d-573">String</span><span class="sxs-lookup"><span data-stu-id="0029d-573">String</span></span>||<span data-ttu-id="0029d-574">EWS 要求です。</span><span class="sxs-lookup"><span data-stu-id="0029d-574">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="0029d-575">function</span><span class="sxs-lookup"><span data-stu-id="0029d-575">function</span></span>||<span data-ttu-id="0029d-576">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="0029d-576">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="0029d-p142">The XML result of the EWS call is provided as a string in the `asyncResult.value` property. If the result exceeds 1 MB in size, an error message is returned instead.</span><span class="sxs-lookup"><span data-stu-id="0029d-p142">The XML result of the EWS call is provided as a string in the `asyncResult.value` property. If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="0029d-579">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="0029d-579">Object</span></span>| <span data-ttu-id="0029d-580">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="0029d-580">&lt;optional&gt;</span></span>|<span data-ttu-id="0029d-581">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="0029d-581">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0029d-582">要件</span><span class="sxs-lookup"><span data-stu-id="0029d-582">Requirements</span></span>

|<span data-ttu-id="0029d-583">要件</span><span class="sxs-lookup"><span data-stu-id="0029d-583">Requirement</span></span>| <span data-ttu-id="0029d-584">値</span><span class="sxs-lookup"><span data-stu-id="0029d-584">Value</span></span>|
|---|---|
|[<span data-ttu-id="0029d-585">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0029d-585">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0029d-586">1.0</span><span class="sxs-lookup"><span data-stu-id="0029d-586">1.0</span></span>|
|[<span data-ttu-id="0029d-587">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0029d-587">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0029d-588">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="0029d-588">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="0029d-589">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0029d-589">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0029d-590">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0029d-590">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0029d-591">例</span><span class="sxs-lookup"><span data-stu-id="0029d-591">Example</span></span>

<span data-ttu-id="0029d-592">次の例は、`GetItem` 操作を使ってアイテムの件名を取得するため、`makeEwsRequestAsync` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="0029d-592">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="0029d-593">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="0029d-593">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="0029d-594">サポートされているイベントの種類のイベント ハンドラーを削除します。</span><span class="sxs-lookup"><span data-stu-id="0029d-594">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="0029d-595">現在、サポートされているイベントの種類は `Office.EventType.ItemChanged` だけです。</span><span class="sxs-lookup"><span data-stu-id="0029d-595">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0029d-596">パラメーター</span><span class="sxs-lookup"><span data-stu-id="0029d-596">Parameters</span></span>

| <span data-ttu-id="0029d-597">名前</span><span class="sxs-lookup"><span data-stu-id="0029d-597">Name</span></span> | <span data-ttu-id="0029d-598">型</span><span class="sxs-lookup"><span data-stu-id="0029d-598">Type</span></span> | <span data-ttu-id="0029d-599">属性</span><span class="sxs-lookup"><span data-stu-id="0029d-599">Attributes</span></span> | <span data-ttu-id="0029d-600">説明</span><span class="sxs-lookup"><span data-stu-id="0029d-600">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="0029d-601">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="0029d-601">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="0029d-602">ハンドラーを取り消すイベント。</span><span class="sxs-lookup"><span data-stu-id="0029d-602">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="0029d-603">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="0029d-603">Object</span></span> | <span data-ttu-id="0029d-604">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0029d-604">&lt;optional&gt;</span></span> | <span data-ttu-id="0029d-605">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="0029d-605">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="0029d-606">Object</span><span class="sxs-lookup"><span data-stu-id="0029d-606">Object</span></span> | <span data-ttu-id="0029d-607">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0029d-607">&lt;optional&gt;</span></span> | <span data-ttu-id="0029d-608">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="0029d-608">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="0029d-609">function</span><span class="sxs-lookup"><span data-stu-id="0029d-609">function</span></span>| <span data-ttu-id="0029d-610">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0029d-610">&lt;optional&gt;</span></span>|<span data-ttu-id="0029d-611">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="0029d-611">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0029d-612">要件</span><span class="sxs-lookup"><span data-stu-id="0029d-612">Requirements</span></span>

|<span data-ttu-id="0029d-613">要件</span><span class="sxs-lookup"><span data-stu-id="0029d-613">Requirement</span></span>| <span data-ttu-id="0029d-614">値</span><span class="sxs-lookup"><span data-stu-id="0029d-614">Value</span></span>|
|---|---|
|[<span data-ttu-id="0029d-615">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0029d-615">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0029d-616">1.5</span><span class="sxs-lookup"><span data-stu-id="0029d-616">1.5</span></span> |
|[<span data-ttu-id="0029d-617">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0029d-617">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0029d-618">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0029d-618">ReadItem</span></span> |
|[<span data-ttu-id="0029d-619">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0029d-619">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0029d-620">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0029d-620">Compose or Read</span></span>|
