---
title: Office.context.mailbox - 要件セット 1.5
description: ''
ms.date: 10/21/2019
localization_priority: Priority
ms.openlocfilehash: bb63d8186d41d072aa62b180b16958d61ce9a66c
ms.sourcegitcommit: 499bf49b41205f8034c501d4db5fe4b02dab205e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/22/2019
ms.locfileid: "37627014"
---
# <a name="mailbox"></a><span data-ttu-id="ea709-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="ea709-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="ea709-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="ea709-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="ea709-104">Microsoft Outlook の Outlook アドイン オブジェクト モデルへのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="ea709-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="ea709-105">要件</span><span class="sxs-lookup"><span data-stu-id="ea709-105">Requirements</span></span>

|<span data-ttu-id="ea709-106">要件</span><span class="sxs-lookup"><span data-stu-id="ea709-106">Requirement</span></span>| <span data-ttu-id="ea709-107">値</span><span class="sxs-lookup"><span data-stu-id="ea709-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="ea709-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ea709-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ea709-109">1.0</span><span class="sxs-lookup"><span data-stu-id="ea709-109">1.0</span></span>|
|[<span data-ttu-id="ea709-110">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ea709-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ea709-111">制限あり</span><span class="sxs-lookup"><span data-stu-id="ea709-111">Restricted</span></span>|
|[<span data-ttu-id="ea709-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ea709-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ea709-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ea709-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="ea709-114">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="ea709-114">Members and methods</span></span>

| <span data-ttu-id="ea709-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="ea709-115">Member</span></span> | <span data-ttu-id="ea709-116">種類</span><span class="sxs-lookup"><span data-stu-id="ea709-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="ea709-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="ea709-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="ea709-118">メンバー</span><span class="sxs-lookup"><span data-stu-id="ea709-118">Member</span></span> |
| [<span data-ttu-id="ea709-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="ea709-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="ea709-120">メンバー</span><span class="sxs-lookup"><span data-stu-id="ea709-120">Member</span></span> |
| [<span data-ttu-id="ea709-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="ea709-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="ea709-122">メソッド</span><span class="sxs-lookup"><span data-stu-id="ea709-122">Method</span></span> |
| [<span data-ttu-id="ea709-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="ea709-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="ea709-124">メソッド</span><span class="sxs-lookup"><span data-stu-id="ea709-124">Method</span></span> |
| [<span data-ttu-id="ea709-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="ea709-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="ea709-126">メソッド</span><span class="sxs-lookup"><span data-stu-id="ea709-126">Method</span></span> |
| [<span data-ttu-id="ea709-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="ea709-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="ea709-128">メソッド</span><span class="sxs-lookup"><span data-stu-id="ea709-128">Method</span></span> |
| [<span data-ttu-id="ea709-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="ea709-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="ea709-130">メソッド</span><span class="sxs-lookup"><span data-stu-id="ea709-130">Method</span></span> |
| [<span data-ttu-id="ea709-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="ea709-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="ea709-132">メソッド</span><span class="sxs-lookup"><span data-stu-id="ea709-132">Method</span></span> |
| [<span data-ttu-id="ea709-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="ea709-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="ea709-134">メソッド</span><span class="sxs-lookup"><span data-stu-id="ea709-134">Method</span></span> |
| [<span data-ttu-id="ea709-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="ea709-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="ea709-136">メソッド</span><span class="sxs-lookup"><span data-stu-id="ea709-136">Method</span></span> |
| [<span data-ttu-id="ea709-137">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="ea709-137">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="ea709-138">メソッド</span><span class="sxs-lookup"><span data-stu-id="ea709-138">Method</span></span> |
| [<span data-ttu-id="ea709-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="ea709-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="ea709-140">メソッド</span><span class="sxs-lookup"><span data-stu-id="ea709-140">Method</span></span> |
| [<span data-ttu-id="ea709-141">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="ea709-141">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="ea709-142">メソッド</span><span class="sxs-lookup"><span data-stu-id="ea709-142">Method</span></span> |
| [<span data-ttu-id="ea709-143">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="ea709-143">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="ea709-144">メソッド</span><span class="sxs-lookup"><span data-stu-id="ea709-144">Method</span></span> |
| [<span data-ttu-id="ea709-145">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="ea709-145">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="ea709-146">メソッド</span><span class="sxs-lookup"><span data-stu-id="ea709-146">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="ea709-147">名前空間</span><span class="sxs-lookup"><span data-stu-id="ea709-147">Namespaces</span></span>

<span data-ttu-id="ea709-148">[diagnostics](Office.context.mailbox.diagnostics.md):Outlook アドインに診断情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="ea709-148">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="ea709-149">[item](Office.context.mailbox.item.md):Outlook アドインのメッセージや予定にアクセスするためのメソッドとプロパティを提供します。</span><span class="sxs-lookup"><span data-stu-id="ea709-149">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="ea709-150">[userProfile](Office.context.mailbox.userProfile.md):Outlook アドインのユーザーに関する情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="ea709-150">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="ea709-151">Members</span><span class="sxs-lookup"><span data-stu-id="ea709-151">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="ea709-152">ewsUrl: String</span><span class="sxs-lookup"><span data-stu-id="ea709-152">ewsUrl: String</span></span>

<span data-ttu-id="ea709-p101">このメール アカウントの Exchange Web サービス (EWS) エンドポイントの URL を取得します。読み取りモードのみです。</span><span class="sxs-lookup"><span data-stu-id="ea709-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="ea709-155">このメンバーは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="ea709-155">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="ea709-p102">
  `ewsUrl\` 値は、リモート サービスで、ユーザーのメールボックスに EWS 呼び出しを行うために使うことができます。たとえば、[選択したアイテムから添付ファイルを取得する](/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="ea709-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="ea709-158">アプリが閲覧モードで `ewsUrl` メンバーを呼び出すには、アプリのマニフェスト内に **ReadItem** アクセス許可が指定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="ea709-158">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="ea709-p103">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出してから、`ewsUrl` メンバーを使用する必要があります。アプリには、`saveAsync` メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="ea709-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="ea709-161">型</span><span class="sxs-lookup"><span data-stu-id="ea709-161">Type</span></span>

*   <span data-ttu-id="ea709-162">String</span><span class="sxs-lookup"><span data-stu-id="ea709-162">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ea709-163">要件</span><span class="sxs-lookup"><span data-stu-id="ea709-163">Requirements</span></span>

|<span data-ttu-id="ea709-164">要件</span><span class="sxs-lookup"><span data-stu-id="ea709-164">Requirement</span></span>| <span data-ttu-id="ea709-165">値</span><span class="sxs-lookup"><span data-stu-id="ea709-165">Value</span></span>|
|---|---|
|[<span data-ttu-id="ea709-166">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ea709-166">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ea709-167">1.0</span><span class="sxs-lookup"><span data-stu-id="ea709-167">1.0</span></span>|
|[<span data-ttu-id="ea709-168">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ea709-168">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ea709-169">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ea709-169">ReadItem</span></span>|
|[<span data-ttu-id="ea709-170">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ea709-170">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ea709-171">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ea709-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="resturl-string"></a><span data-ttu-id="ea709-172">restUrl: String</span><span class="sxs-lookup"><span data-stu-id="ea709-172">restUrl: String</span></span>

<span data-ttu-id="ea709-173">この電子メール アカウントの REST エンドポイントの URL を取得します。</span><span class="sxs-lookup"><span data-stu-id="ea709-173">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="ea709-174">`restUrl` 値は、ユーザーのメールボックスに [REST API](/outlook/rest/) 呼び出しを行うために使用できます。</span><span class="sxs-lookup"><span data-stu-id="ea709-174">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="ea709-175">アプリが閲覧モードで `restUrl` メンバーを呼び出すには、アプリのマニフェスト内に **ReadItem** アクセス許可が指定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="ea709-175">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="ea709-p104">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出してから、`restUrl` メンバーを使用する必要があります。アプリには、`saveAsync` メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="ea709-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

> [!NOTE]
> <span data-ttu-id="ea709-178">構成されたカスタム REST URL を使用する Exchange 2016 以降のオンプレミスのインストールに接続されている Outlook クライアントは、`restUrl` に無効な値を返します。</span><span class="sxs-lookup"><span data-stu-id="ea709-178">Outlook clients connected to on-premises installations of Exchange 2016 or later with a custom REST URL configured will return an invalid value for `restUrl`.</span></span>

##### <a name="type"></a><span data-ttu-id="ea709-179">型</span><span class="sxs-lookup"><span data-stu-id="ea709-179">Type</span></span>

*   <span data-ttu-id="ea709-180">String</span><span class="sxs-lookup"><span data-stu-id="ea709-180">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ea709-181">要件</span><span class="sxs-lookup"><span data-stu-id="ea709-181">Requirements</span></span>

|<span data-ttu-id="ea709-182">要件</span><span class="sxs-lookup"><span data-stu-id="ea709-182">Requirement</span></span>| <span data-ttu-id="ea709-183">値</span><span class="sxs-lookup"><span data-stu-id="ea709-183">Value</span></span>|
|---|---|
|[<span data-ttu-id="ea709-184">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ea709-184">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ea709-185">1.5</span><span class="sxs-lookup"><span data-stu-id="ea709-185">1.5</span></span> |
|[<span data-ttu-id="ea709-186">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ea709-186">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ea709-187">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ea709-187">ReadItem</span></span>|
|[<span data-ttu-id="ea709-188">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ea709-188">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ea709-189">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ea709-189">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="ea709-190">メソッド</span><span class="sxs-lookup"><span data-stu-id="ea709-190">Methods</span></span>

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="ea709-191">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="ea709-191">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="ea709-192">サポートされているイベントのイベント ハンドラーを追加します。</span><span class="sxs-lookup"><span data-stu-id="ea709-192">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="ea709-193">現在サポートされている唯一のイベントの種類は `Office.EventType.ItemChanged` です。これはユーザーが新しいアイテムを選択したときに呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="ea709-193">Currently the only supported event type is `Office.EventType.ItemChanged`, which is invoked when the user selects a new item.</span></span> <span data-ttu-id="ea709-194">このイベントは、ピン留め可能な作業ウィンドウを実装するアドインで使用され、現在選択されているアイテムに基づいて作業ウィンドウ UI をアドインで更新できるようにします。</span><span class="sxs-lookup"><span data-stu-id="ea709-194">This event is used by add-ins that implement a pinnable task pane, and allows the add-in to refresh the task pane UI based on the currently selected item.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ea709-195">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ea709-195">Parameters</span></span>

| <span data-ttu-id="ea709-196">名前</span><span class="sxs-lookup"><span data-stu-id="ea709-196">Name</span></span> | <span data-ttu-id="ea709-197">型</span><span class="sxs-lookup"><span data-stu-id="ea709-197">Type</span></span> | <span data-ttu-id="ea709-198">属性</span><span class="sxs-lookup"><span data-stu-id="ea709-198">Attributes</span></span> | <span data-ttu-id="ea709-199">説明</span><span class="sxs-lookup"><span data-stu-id="ea709-199">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="ea709-200">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="ea709-200">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="ea709-201">ハンドラーを呼び出す必要のあるイベント。</span><span class="sxs-lookup"><span data-stu-id="ea709-201">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="ea709-202">Function</span><span class="sxs-lookup"><span data-stu-id="ea709-202">Function</span></span> || <span data-ttu-id="ea709-p106">イベントを処理する関数。関数は、オブジェクト リテラルである単一パラメーターを受け入れる必要があります。パラメーターの `type` プロパティは、`addHandlerAsync` に渡される `eventType` パラメーターと一致します。</span><span class="sxs-lookup"><span data-stu-id="ea709-p106">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="ea709-206">Object</span><span class="sxs-lookup"><span data-stu-id="ea709-206">Object</span></span> | <span data-ttu-id="ea709-207">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ea709-207">&lt;optional&gt;</span></span> | <span data-ttu-id="ea709-208">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="ea709-208">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="ea709-209">Object</span><span class="sxs-lookup"><span data-stu-id="ea709-209">Object</span></span> | <span data-ttu-id="ea709-210">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ea709-210">&lt;optional&gt;</span></span> | <span data-ttu-id="ea709-211">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="ea709-211">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="ea709-212">function</span><span class="sxs-lookup"><span data-stu-id="ea709-212">function</span></span>| <span data-ttu-id="ea709-213">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ea709-213">&lt;optional&gt;</span></span>|<span data-ttu-id="ea709-214">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="ea709-214">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ea709-215">要件</span><span class="sxs-lookup"><span data-stu-id="ea709-215">Requirements</span></span>

|<span data-ttu-id="ea709-216">要件</span><span class="sxs-lookup"><span data-stu-id="ea709-216">Requirement</span></span>| <span data-ttu-id="ea709-217">値</span><span class="sxs-lookup"><span data-stu-id="ea709-217">Value</span></span>|
|---|---|
|[<span data-ttu-id="ea709-218">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ea709-218">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ea709-219">1.5</span><span class="sxs-lookup"><span data-stu-id="ea709-219">1.5</span></span> |
|[<span data-ttu-id="ea709-220">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ea709-220">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ea709-221">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ea709-221">ReadItem</span></span> |
|[<span data-ttu-id="ea709-222">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ea709-222">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ea709-223">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ea709-223">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ea709-224">例</span><span class="sxs-lookup"><span data-stu-id="ea709-224">Example</span></span>

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

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="ea709-225">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="ea709-225">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="ea709-226">REST 形式のアイテム ID を EWS 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="ea709-226">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="ea709-227">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="ea709-227">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="ea709-p107">REST API ([Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) や [Microsoft Graph](https://graph.microsoft.io/) など) で取得されたアイテム ID は、Exchange Web サービス (EWS) に使用される形式とは異なる形式を使用します。`convertToEwsId` メソッドは、REST 形式の ID を EWS 用の適切な形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="ea709-p107">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ea709-230">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ea709-230">Parameters</span></span>

|<span data-ttu-id="ea709-231">名前</span><span class="sxs-lookup"><span data-stu-id="ea709-231">Name</span></span>| <span data-ttu-id="ea709-232">種類</span><span class="sxs-lookup"><span data-stu-id="ea709-232">Type</span></span>| <span data-ttu-id="ea709-233">説明</span><span class="sxs-lookup"><span data-stu-id="ea709-233">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="ea709-234">String</span><span class="sxs-lookup"><span data-stu-id="ea709-234">String</span></span>|<span data-ttu-id="ea709-235">Outlook REST API 形式のアイテム ID</span><span class="sxs-lookup"><span data-stu-id="ea709-235">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="ea709-236">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="ea709-236">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.5)|<span data-ttu-id="ea709-237">アイテム ID の取得に使用された Outlook REST API のバージョンを示す値。</span><span class="sxs-lookup"><span data-stu-id="ea709-237">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ea709-238">要件</span><span class="sxs-lookup"><span data-stu-id="ea709-238">Requirements</span></span>

|<span data-ttu-id="ea709-239">要件</span><span class="sxs-lookup"><span data-stu-id="ea709-239">Requirement</span></span>| <span data-ttu-id="ea709-240">値</span><span class="sxs-lookup"><span data-stu-id="ea709-240">Value</span></span>|
|---|---|
|[<span data-ttu-id="ea709-241">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ea709-241">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ea709-242">1.3</span><span class="sxs-lookup"><span data-stu-id="ea709-242">1.3</span></span>|
|[<span data-ttu-id="ea709-243">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ea709-243">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ea709-244">制限あり</span><span class="sxs-lookup"><span data-stu-id="ea709-244">Restricted</span></span>|
|[<span data-ttu-id="ea709-245">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ea709-245">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ea709-246">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ea709-246">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ea709-247">戻り値:</span><span class="sxs-lookup"><span data-stu-id="ea709-247">Returns:</span></span>

<span data-ttu-id="ea709-248">型:String</span><span class="sxs-lookup"><span data-stu-id="ea709-248">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="ea709-249">例</span><span class="sxs-lookup"><span data-stu-id="ea709-249">Example</span></span>

```js
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-15"></a><span data-ttu-id="ea709-250">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.5)}</span><span class="sxs-lookup"><span data-stu-id="ea709-250">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.5)}</span></span>

<span data-ttu-id="ea709-251">クライアントのローカル時間で時間情報が含まれている辞書を取得します。</span><span class="sxs-lookup"><span data-stu-id="ea709-251">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="ea709-p108">Outlook on the web または Outlook デスクトップのメールアプリでは、日付と時刻に異なるタイムゾーンを使用できます。Outlook デスクトップは、クライアント コンピューターのタイム ゾーンを使用します。Outlook on the web は、Exchange 管理センター (EAC) で設定されたタイム ゾーンを使用します。日付と時刻の値は、ユーザー インターフェイスに表示される値が、常にユーザーが期待するタイム ゾーンと一致するように処理する必要があります。</span><span class="sxs-lookup"><span data-stu-id="ea709-p108">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times. Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="ea709-p109">Outlook デスクトップ クライアントでメール アプリを実行している場合、`convertToLocalClientTime` メソッドは、クライアント コンピューターのタイム ゾーンに設定された値のディクショナリ オブジェクトを返します。Outlook on the web でメール アプリを実行している場合、`convertToLocalClientTime` メソッドは、EAC で指定したタイム ゾーンに設定された値のディクショナリ オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="ea709-p109">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ea709-257">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ea709-257">Parameters</span></span>

|<span data-ttu-id="ea709-258">名前</span><span class="sxs-lookup"><span data-stu-id="ea709-258">Name</span></span>| <span data-ttu-id="ea709-259">種類</span><span class="sxs-lookup"><span data-stu-id="ea709-259">Type</span></span>| <span data-ttu-id="ea709-260">説明</span><span class="sxs-lookup"><span data-stu-id="ea709-260">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="ea709-261">Date</span><span class="sxs-lookup"><span data-stu-id="ea709-261">Date</span></span>|<span data-ttu-id="ea709-262">日付オブジェクト</span><span class="sxs-lookup"><span data-stu-id="ea709-262">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ea709-263">要件</span><span class="sxs-lookup"><span data-stu-id="ea709-263">Requirements</span></span>

|<span data-ttu-id="ea709-264">要件</span><span class="sxs-lookup"><span data-stu-id="ea709-264">Requirement</span></span>| <span data-ttu-id="ea709-265">値</span><span class="sxs-lookup"><span data-stu-id="ea709-265">Value</span></span>|
|---|---|
|[<span data-ttu-id="ea709-266">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ea709-266">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ea709-267">1.0</span><span class="sxs-lookup"><span data-stu-id="ea709-267">1.0</span></span>|
|[<span data-ttu-id="ea709-268">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ea709-268">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ea709-269">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ea709-269">ReadItem</span></span>|
|[<span data-ttu-id="ea709-270">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ea709-270">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ea709-271">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ea709-271">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ea709-272">戻り値:</span><span class="sxs-lookup"><span data-stu-id="ea709-272">Returns:</span></span>

<span data-ttu-id="ea709-273">型:[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="ea709-273">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.5)</span></span>

<br>

---
---

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="ea709-274">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="ea709-274">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="ea709-275">EWS 形式のアイテム ID を REST 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="ea709-275">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="ea709-276">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="ea709-276">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="ea709-p110">EWS または `itemId` プロパティで取得されるアイテム ID は、REST API ([Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) や [Microsoft Graph](https://graph.microsoft.io/) など) に使用される形式とは異なる形式を使用します。`convertToRestId` メソッドは、EWS 形式の ID を REST 用の適切な形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="ea709-p110">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ea709-279">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ea709-279">Parameters</span></span>

|<span data-ttu-id="ea709-280">名前</span><span class="sxs-lookup"><span data-stu-id="ea709-280">Name</span></span>| <span data-ttu-id="ea709-281">種類</span><span class="sxs-lookup"><span data-stu-id="ea709-281">Type</span></span>| <span data-ttu-id="ea709-282">説明</span><span class="sxs-lookup"><span data-stu-id="ea709-282">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="ea709-283">String</span><span class="sxs-lookup"><span data-stu-id="ea709-283">String</span></span>|<span data-ttu-id="ea709-284">Exchange Web サービス (EWS) 形式のアイテム ID</span><span class="sxs-lookup"><span data-stu-id="ea709-284">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="ea709-285">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="ea709-285">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.5)|<span data-ttu-id="ea709-286">変換後の ID を使用する Outlook REST API のバージョンを示す値。</span><span class="sxs-lookup"><span data-stu-id="ea709-286">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ea709-287">要件</span><span class="sxs-lookup"><span data-stu-id="ea709-287">Requirements</span></span>

|<span data-ttu-id="ea709-288">要件</span><span class="sxs-lookup"><span data-stu-id="ea709-288">Requirement</span></span>| <span data-ttu-id="ea709-289">値</span><span class="sxs-lookup"><span data-stu-id="ea709-289">Value</span></span>|
|---|---|
|[<span data-ttu-id="ea709-290">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ea709-290">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ea709-291">1.3</span><span class="sxs-lookup"><span data-stu-id="ea709-291">1.3</span></span>|
|[<span data-ttu-id="ea709-292">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ea709-292">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ea709-293">制限あり</span><span class="sxs-lookup"><span data-stu-id="ea709-293">Restricted</span></span>|
|[<span data-ttu-id="ea709-294">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ea709-294">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ea709-295">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ea709-295">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ea709-296">戻り値:</span><span class="sxs-lookup"><span data-stu-id="ea709-296">Returns:</span></span>

<span data-ttu-id="ea709-297">型:String</span><span class="sxs-lookup"><span data-stu-id="ea709-297">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="ea709-298">例</span><span class="sxs-lookup"><span data-stu-id="ea709-298">Example</span></span>

```js
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="ea709-299">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="ea709-299">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="ea709-300">時間情報が含まれているディクショナリから日付オブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="ea709-300">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="ea709-301">`convertToUtcClientTime` メソッドは、ローカルの日付と時刻を含むディクショナリを、ローカルの日付と時刻の正しい値を持つ日付オブジェクトに変換します。</span><span class="sxs-lookup"><span data-stu-id="ea709-301">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ea709-302">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ea709-302">Parameters</span></span>

|<span data-ttu-id="ea709-303">名前</span><span class="sxs-lookup"><span data-stu-id="ea709-303">Name</span></span>| <span data-ttu-id="ea709-304">種類</span><span class="sxs-lookup"><span data-stu-id="ea709-304">Type</span></span>| <span data-ttu-id="ea709-305">説明</span><span class="sxs-lookup"><span data-stu-id="ea709-305">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="ea709-306">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="ea709-306">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.5)|<span data-ttu-id="ea709-307">変換するローカル時刻の値。</span><span class="sxs-lookup"><span data-stu-id="ea709-307">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ea709-308">要件</span><span class="sxs-lookup"><span data-stu-id="ea709-308">Requirements</span></span>

|<span data-ttu-id="ea709-309">要件</span><span class="sxs-lookup"><span data-stu-id="ea709-309">Requirement</span></span>| <span data-ttu-id="ea709-310">値</span><span class="sxs-lookup"><span data-stu-id="ea709-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="ea709-311">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ea709-311">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ea709-312">1.0</span><span class="sxs-lookup"><span data-stu-id="ea709-312">1.0</span></span>|
|[<span data-ttu-id="ea709-313">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ea709-313">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ea709-314">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ea709-314">ReadItem</span></span>|
|[<span data-ttu-id="ea709-315">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ea709-315">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ea709-316">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ea709-316">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ea709-317">戻り値:</span><span class="sxs-lookup"><span data-stu-id="ea709-317">Returns:</span></span>

<span data-ttu-id="ea709-318">時間が UTC で表現された日付オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="ea709-318">A Date object with the time expressed in UTC.</span></span>

<span data-ttu-id="ea709-319">型: Date</span><span class="sxs-lookup"><span data-stu-id="ea709-319">Type: Date</span></span>

##### <a name="example"></a><span data-ttu-id="ea709-320">例</span><span class="sxs-lookup"><span data-stu-id="ea709-320">Example</span></span>

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

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="ea709-321">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="ea709-321">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="ea709-322">既存の予定を表示します。</span><span class="sxs-lookup"><span data-stu-id="ea709-322">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="ea709-323">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="ea709-323">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="ea709-324">`displayAppointmentForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存の予定を開きます。</span><span class="sxs-lookup"><span data-stu-id="ea709-324">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="ea709-p111">Outlook on Mac では、このメソッドを使用して定期的な系列に含まれない単発の予定や定期的な系列のマスター予定を表示できます。ただし、系列のインスタンスは表示できません。これは、Outlook on Mac では定期的な系列のインスタンスのプロパティ (アイテム ID を含む) にアクセスできないためです。</span><span class="sxs-lookup"><span data-stu-id="ea709-p111">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="ea709-327">Outlook on the web では、このメソッドはフォームの本文が 32KB 以下の文字数の場合にのみ指定のフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="ea709-327">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="ea709-328">指定のアイテム識別子が既存の予定を表していない場合は、クライアント コンピューターまたはデバイスで空のウィンドウが開き、エラー メッセージは返されません。</span><span class="sxs-lookup"><span data-stu-id="ea709-328">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ea709-329">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ea709-329">Parameters</span></span>

|<span data-ttu-id="ea709-330">名前</span><span class="sxs-lookup"><span data-stu-id="ea709-330">Name</span></span>| <span data-ttu-id="ea709-331">種類</span><span class="sxs-lookup"><span data-stu-id="ea709-331">Type</span></span>| <span data-ttu-id="ea709-332">説明</span><span class="sxs-lookup"><span data-stu-id="ea709-332">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="ea709-333">String</span><span class="sxs-lookup"><span data-stu-id="ea709-333">String</span></span>|<span data-ttu-id="ea709-334">既存の予定の Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="ea709-334">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ea709-335">要件</span><span class="sxs-lookup"><span data-stu-id="ea709-335">Requirements</span></span>

|<span data-ttu-id="ea709-336">要件</span><span class="sxs-lookup"><span data-stu-id="ea709-336">Requirement</span></span>| <span data-ttu-id="ea709-337">値</span><span class="sxs-lookup"><span data-stu-id="ea709-337">Value</span></span>|
|---|---|
|[<span data-ttu-id="ea709-338">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ea709-338">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ea709-339">1.0</span><span class="sxs-lookup"><span data-stu-id="ea709-339">1.0</span></span>|
|[<span data-ttu-id="ea709-340">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ea709-340">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ea709-341">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ea709-341">ReadItem</span></span>|
|[<span data-ttu-id="ea709-342">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ea709-342">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ea709-343">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ea709-343">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ea709-344">例</span><span class="sxs-lookup"><span data-stu-id="ea709-344">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

<br>

---
---

#### <a name="displaymessageformitemid"></a><span data-ttu-id="ea709-345">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="ea709-345">displayMessageForm(itemId)</span></span>

<span data-ttu-id="ea709-346">既存のメッセージを表示します。</span><span class="sxs-lookup"><span data-stu-id="ea709-346">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="ea709-347">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="ea709-347">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="ea709-348">`displayMessageForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存のメッセージを開きます。</span><span class="sxs-lookup"><span data-stu-id="ea709-348">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="ea709-349">Outlook on the web では、このメソッドはフォームの本文が 32 KB 以下の文字数の場合にのみ指定のフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="ea709-349">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="ea709-350">指定のアイテム識別子が既存のメッセージを表していない場合は、クライアント コンピューターにはメッセージは表示されず、エラー メッセージも返されません。</span><span class="sxs-lookup"><span data-stu-id="ea709-350">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="ea709-p112">予定を表す `itemId` を含む `displayMessageForm` を使用しないでください。既存の予定を表示するには、`displayAppointmentForm` メソッドを使用します。新しい予定を作成するフォームを表示するには、`displayNewAppointmentForm` メソッドを使用します。</span><span class="sxs-lookup"><span data-stu-id="ea709-p112">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ea709-353">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ea709-353">Parameters</span></span>

|<span data-ttu-id="ea709-354">名前</span><span class="sxs-lookup"><span data-stu-id="ea709-354">Name</span></span>| <span data-ttu-id="ea709-355">種類</span><span class="sxs-lookup"><span data-stu-id="ea709-355">Type</span></span>| <span data-ttu-id="ea709-356">説明</span><span class="sxs-lookup"><span data-stu-id="ea709-356">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="ea709-357">String</span><span class="sxs-lookup"><span data-stu-id="ea709-357">String</span></span>|<span data-ttu-id="ea709-358">既存のメッセージの Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="ea709-358">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ea709-359">要件</span><span class="sxs-lookup"><span data-stu-id="ea709-359">Requirements</span></span>

|<span data-ttu-id="ea709-360">要件</span><span class="sxs-lookup"><span data-stu-id="ea709-360">Requirement</span></span>| <span data-ttu-id="ea709-361">値</span><span class="sxs-lookup"><span data-stu-id="ea709-361">Value</span></span>|
|---|---|
|[<span data-ttu-id="ea709-362">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ea709-362">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ea709-363">1.0</span><span class="sxs-lookup"><span data-stu-id="ea709-363">1.0</span></span>|
|[<span data-ttu-id="ea709-364">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ea709-364">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ea709-365">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ea709-365">ReadItem</span></span>|
|[<span data-ttu-id="ea709-366">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ea709-366">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ea709-367">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ea709-367">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ea709-368">例</span><span class="sxs-lookup"><span data-stu-id="ea709-368">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

<br>

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="ea709-369">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="ea709-369">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="ea709-370">新しい予定を作成するためのフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="ea709-370">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="ea709-371">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="ea709-371">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="ea709-p113">`displayNewAppointmentForm` メソッドを使用すると、ユーザーが新しい予定または会議を作成できるフォームが開きます。パラメーターを指定すると、予定のフォーム フィールドにパラメーターの内容が自動的に設定されます。</span><span class="sxs-lookup"><span data-stu-id="ea709-p113">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="ea709-p114">Outlook on the web およびモバイル デバイスでは、このメソッドは常に出席者フィールドが含まれるフォームを表示します。入力引数として出席者を指定しないと、このメソッドは **[保存]** ボタンのあるフォームを表示します。出席者を指定した場合には、フォームにその出席者と **[送信]** ボタンが表示されます。</span><span class="sxs-lookup"><span data-stu-id="ea709-p114">In Outlook on the web and mobile devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="ea709-p115">Outlook リッチ クライアントと Outlook RT で、`requiredAttendees`、`optionalAttendees`、または `resources` パラメーターに出席者またはリソースを指定し、このメソッドを実行すると、**[送信]** ボタンがある会議フォームが表示されます。受信者を指定せずにこのメソッドを実行すると、**[保存して閉じる]** ボタンがある予定フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="ea709-p115">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="ea709-379">パラメーターのいずれかが指定のサイズ制限を超える場合、または不明なパラメーター名が指定されている場合は、例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="ea709-379">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ea709-380">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ea709-380">Parameters</span></span>

|<span data-ttu-id="ea709-381">名前</span><span class="sxs-lookup"><span data-stu-id="ea709-381">Name</span></span>| <span data-ttu-id="ea709-382">種類</span><span class="sxs-lookup"><span data-stu-id="ea709-382">Type</span></span>| <span data-ttu-id="ea709-383">説明</span><span class="sxs-lookup"><span data-stu-id="ea709-383">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="ea709-384">Object</span><span class="sxs-lookup"><span data-stu-id="ea709-384">Object</span></span> | <span data-ttu-id="ea709-385">新しい予定を記述するパラメーターのディクショナリ。</span><span class="sxs-lookup"><span data-stu-id="ea709-385">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="ea709-386">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)&gt;</span><span class="sxs-lookup"><span data-stu-id="ea709-386">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)&gt;</span></span> | <span data-ttu-id="ea709-p116">予定に必要な各出席者について、メール アドレスを含む文字列の配列、または `EmailAddressDetails` オブジェクトを含む配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="ea709-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="ea709-389">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)&gt;</span><span class="sxs-lookup"><span data-stu-id="ea709-389">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)&gt;</span></span> | <span data-ttu-id="ea709-p117">予定の各任意出席者について、メール アドレスを含む文字列の配列、または `EmailAddressDetails` オブジェクトを含む配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="ea709-p117">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="ea709-392">日付</span><span class="sxs-lookup"><span data-stu-id="ea709-392">Date</span></span> | <span data-ttu-id="ea709-393">予定の開始日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="ea709-393">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="ea709-394">Date</span><span class="sxs-lookup"><span data-stu-id="ea709-394">Date</span></span> | <span data-ttu-id="ea709-395">予定の終了日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="ea709-395">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="ea709-396">String</span><span class="sxs-lookup"><span data-stu-id="ea709-396">String</span></span> | <span data-ttu-id="ea709-p118">予定の場所を含む文字列。文字列は最大 255 文字に制限されます。</span><span class="sxs-lookup"><span data-stu-id="ea709-p118">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="ea709-399">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="ea709-399">Array.&lt;String&gt;</span></span> | <span data-ttu-id="ea709-p119">予定に必要なリソースを含む文字列の配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="ea709-p119">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="ea709-402">String</span><span class="sxs-lookup"><span data-stu-id="ea709-402">String</span></span> | <span data-ttu-id="ea709-p120">予定の件名を含む文字列です。文字列は最大 255 文字に制限されます。</span><span class="sxs-lookup"><span data-stu-id="ea709-p120">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="ea709-405">String</span><span class="sxs-lookup"><span data-stu-id="ea709-405">String</span></span> | <span data-ttu-id="ea709-p121">予定の本文。本文の内容は、最大サイズが 32 KB に制限されます。</span><span class="sxs-lookup"><span data-stu-id="ea709-p121">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="ea709-408">要件</span><span class="sxs-lookup"><span data-stu-id="ea709-408">Requirements</span></span>

|<span data-ttu-id="ea709-409">要件</span><span class="sxs-lookup"><span data-stu-id="ea709-409">Requirement</span></span>| <span data-ttu-id="ea709-410">値</span><span class="sxs-lookup"><span data-stu-id="ea709-410">Value</span></span>|
|---|---|
|[<span data-ttu-id="ea709-411">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ea709-411">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ea709-412">1.0</span><span class="sxs-lookup"><span data-stu-id="ea709-412">1.0</span></span>|
|[<span data-ttu-id="ea709-413">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ea709-413">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ea709-414">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ea709-414">ReadItem</span></span>|
|[<span data-ttu-id="ea709-415">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ea709-415">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ea709-416">読み取り</span><span class="sxs-lookup"><span data-stu-id="ea709-416">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ea709-417">例</span><span class="sxs-lookup"><span data-stu-id="ea709-417">Example</span></span>

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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="ea709-418">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="ea709-418">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="ea709-419">REST API または Exchange Web サービスを呼び出すために使用するトークンを含む文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="ea709-419">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="ea709-p122">`getCallbackTokenAsync` メソッドは、ユーザーのメールボックスをホストする Exchange Server から不透明なトークンを取得する非同期の呼び出しを行います。コールバック トークンの有効期間は 5 分です。</span><span class="sxs-lookup"><span data-stu-id="ea709-p122">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="ea709-422">可能であれば、アドインでは Exchange Web サービスの代わりに REST API を使用することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="ea709-422">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span>

<span data-ttu-id="ea709-423">閲覧モードで `getCallbackTokenAsync` メソッドを呼び出すには、**ReadItem** の最小限のアクセス許可レベルが必要です。</span><span class="sxs-lookup"><span data-stu-id="ea709-423">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="ea709-424">作成モードで `getCallbackTokenAsync` を呼び出すには、アイテムを保存しておく必要があります。</span><span class="sxs-lookup"><span data-stu-id="ea709-424">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="ea709-425">[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドは、**ReadWriteItem** の最小限のアクセス許可レベルが必要です。</span><span class="sxs-lookup"><span data-stu-id="ea709-425">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

<span data-ttu-id="ea709-426">**REST トークン**</span><span class="sxs-lookup"><span data-stu-id="ea709-426">**REST Tokens**</span></span>

<span data-ttu-id="ea709-p124">REST トークンが要求された場合 (`options.isRest = true`)、結果トークンは Exchange Web サービスの呼び出しを認証するためには機能しません。アドインがマニフェストで [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) アクセス許可を指定していない限り、トークンの範囲は現在のアイテムとその添付ファイルへの読み取り専用アクセスに制限されます。`ReadWriteMailbox` アクセス許可が指定されている場合は、結果トークンは、メールを送信する機能など、メール、カレンダー、連絡先への読み取り/書き込みアクセスを付与します。</span><span class="sxs-lookup"><span data-stu-id="ea709-p124">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="ea709-430">アドインでは、`restUrl` プロパティを使用して、REST API 呼び出しを行うときに使用する正しい URL を決定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="ea709-430">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="ea709-431">**EWS トークン**</span><span class="sxs-lookup"><span data-stu-id="ea709-431">**EWS Tokens**</span></span>

<span data-ttu-id="ea709-p125">EWS トークンが要求された場合 (`options.isRest = false`)、結果トークンは REST API 呼び出しを認証するためには機能しません。トークンの範囲は、現在のアイテムへのアクセスに制限されます。</span><span class="sxs-lookup"><span data-stu-id="ea709-p125">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="ea709-434">アドインでは、`ewsUrl` プロパティを使用して、EWS 呼び出しを行うときに使用する正しい URL を決定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="ea709-434">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

<span data-ttu-id="ea709-435">トークンと、添付ファイル識別子またはアイテム識別子の両方をサードパーティ システムに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="ea709-435">You can pass the token and an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="ea709-436">サードパーティ システムは、トークンをベアラー承認トークンとして使用し、Exchange Web サービス (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) 操作または [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) 操作を呼び出して、添付ファイルまたはアイテムを返します。</span><span class="sxs-lookup"><span data-stu-id="ea709-436">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item.</span></span> <span data-ttu-id="ea709-437">たとえば、[選択したアイテムから添付ファイルを取得する](/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="ea709-437">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

##### <a name="parameters"></a><span data-ttu-id="ea709-438">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ea709-438">Parameters</span></span>

|<span data-ttu-id="ea709-439">名前</span><span class="sxs-lookup"><span data-stu-id="ea709-439">Name</span></span>| <span data-ttu-id="ea709-440">型</span><span class="sxs-lookup"><span data-stu-id="ea709-440">Type</span></span>| <span data-ttu-id="ea709-441">属性</span><span class="sxs-lookup"><span data-stu-id="ea709-441">Attributes</span></span>| <span data-ttu-id="ea709-442">説明</span><span class="sxs-lookup"><span data-stu-id="ea709-442">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="ea709-443">Object</span><span class="sxs-lookup"><span data-stu-id="ea709-443">Object</span></span> | <span data-ttu-id="ea709-444">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ea709-444">&lt;optional&gt;</span></span> | <span data-ttu-id="ea709-445">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="ea709-445">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="ea709-446">Boolean</span><span class="sxs-lookup"><span data-stu-id="ea709-446">Boolean</span></span> |  <span data-ttu-id="ea709-447">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="ea709-447">&lt;optional&gt;</span></span> | <span data-ttu-id="ea709-p127">提供されたトークンを Outlook REST API または Exchange Web サービスに使用するかどうかを決定します。既定値は、`false` です。</span><span class="sxs-lookup"><span data-stu-id="ea709-p127">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="ea709-450">Object</span><span class="sxs-lookup"><span data-stu-id="ea709-450">Object</span></span> |  <span data-ttu-id="ea709-451">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="ea709-451">&lt;optional&gt;</span></span> | <span data-ttu-id="ea709-452">非同期メソッドに渡される状態データ。</span><span class="sxs-lookup"><span data-stu-id="ea709-452">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="ea709-453">function</span><span class="sxs-lookup"><span data-stu-id="ea709-453">function</span></span>||<span data-ttu-id="ea709-454">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="ea709-454">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="ea709-455">トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="ea709-455">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="ea709-456">エラーが発生した場合、 `asyncResult.error` および `asyncResult.diagnostics` のプロパティで追加情報が提供される場合があります。</span><span class="sxs-lookup"><span data-stu-id="ea709-456">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|

##### <a name="errors"></a><span data-ttu-id="ea709-457">エラー</span><span class="sxs-lookup"><span data-stu-id="ea709-457">Errors</span></span>

|<span data-ttu-id="ea709-458">エラー コード</span><span class="sxs-lookup"><span data-stu-id="ea709-458">Error code</span></span>|<span data-ttu-id="ea709-459">説明</span><span class="sxs-lookup"><span data-stu-id="ea709-459">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="ea709-460">要求が失敗しました。</span><span class="sxs-lookup"><span data-stu-id="ea709-460">The request has failed.</span></span> <span data-ttu-id="ea709-461">HTTP エラーコードの diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="ea709-461">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="ea709-462">Exchange サーバーがエラーを返しました。</span><span class="sxs-lookup"><span data-stu-id="ea709-462">The Exchange server returned an error.</span></span> <span data-ttu-id="ea709-463">詳細については、diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="ea709-463">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="ea709-464">ユーザーはネットワークに接続されていません。</span><span class="sxs-lookup"><span data-stu-id="ea709-464">The user is no longer connected to the network.</span></span> <span data-ttu-id="ea709-465">ネットワーク接続を確認し、やり直してください。</span><span class="sxs-lookup"><span data-stu-id="ea709-465">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ea709-466">要件</span><span class="sxs-lookup"><span data-stu-id="ea709-466">Requirements</span></span>

|<span data-ttu-id="ea709-467">要件</span><span class="sxs-lookup"><span data-stu-id="ea709-467">Requirement</span></span>| <span data-ttu-id="ea709-468">値</span><span class="sxs-lookup"><span data-stu-id="ea709-468">Value</span></span>|
|---|---|
|[<span data-ttu-id="ea709-469">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ea709-469">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ea709-470">1.5</span><span class="sxs-lookup"><span data-stu-id="ea709-470">1.5</span></span> |
|[<span data-ttu-id="ea709-471">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ea709-471">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ea709-472">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ea709-472">ReadItem</span></span>|
|[<span data-ttu-id="ea709-473">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ea709-473">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ea709-474">新規作成と閲覧</span><span class="sxs-lookup"><span data-stu-id="ea709-474">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="ea709-475">例</span><span class="sxs-lookup"><span data-stu-id="ea709-475">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="ea709-476">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="ea709-476">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="ea709-477">Exchange Server から添付ファイルやアイテムを取得するために使うトークンを含む文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="ea709-477">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="ea709-p131">`getCallbackTokenAsync` メソッドは、ユーザーのメールボックスをホストする Exchange Server から不透明なトークンを取得する非同期の呼び出しを行います。コールバック トークンの有効期間は 5 分です。</span><span class="sxs-lookup"><span data-stu-id="ea709-p131">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="ea709-480">トークンと、添付ファイル識別子またはアイテム識別子の両方をサードパーティ システムに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="ea709-480">You can pass the token and an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="ea709-481">サードパーティ システムは、トークンをベアラー承認トークンとして使用し、Exchange Web サービス (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) 操作または [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) 操作を呼び出して、添付ファイルまたはアイテムを返します。</span><span class="sxs-lookup"><span data-stu-id="ea709-481">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item.</span></span> <span data-ttu-id="ea709-482">たとえば、[選択したアイテムから添付ファイルを取得する](/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="ea709-482">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="ea709-483">閲覧モードで `getCallbackTokenAsync` メソッドを呼び出すには、**ReadItem** の最小限のアクセス許可レベルが必要です。</span><span class="sxs-lookup"><span data-stu-id="ea709-483">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="ea709-484">作成モードで `getCallbackTokenAsync` を呼び出すには、アイテムを保存しておく必要があります。</span><span class="sxs-lookup"><span data-stu-id="ea709-484">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="ea709-485">[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドは、**ReadWriteItem** の最小限のアクセス許可レベルが必要です。</span><span class="sxs-lookup"><span data-stu-id="ea709-485">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ea709-486">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ea709-486">Parameters</span></span>

|<span data-ttu-id="ea709-487">名前</span><span class="sxs-lookup"><span data-stu-id="ea709-487">Name</span></span>| <span data-ttu-id="ea709-488">型</span><span class="sxs-lookup"><span data-stu-id="ea709-488">Type</span></span>| <span data-ttu-id="ea709-489">属性</span><span class="sxs-lookup"><span data-stu-id="ea709-489">Attributes</span></span>| <span data-ttu-id="ea709-490">説明</span><span class="sxs-lookup"><span data-stu-id="ea709-490">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="ea709-491">function</span><span class="sxs-lookup"><span data-stu-id="ea709-491">function</span></span>||<span data-ttu-id="ea709-492">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="ea709-492">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="ea709-493">トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="ea709-493">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="ea709-494">エラーが発生した場合、 `asyncResult.error` および `asyncResult.diagnostics` のプロパティで追加情報が提供される場合があります。</span><span class="sxs-lookup"><span data-stu-id="ea709-494">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="ea709-495">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="ea709-495">Object</span></span>| <span data-ttu-id="ea709-496">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="ea709-496">&lt;optional&gt;</span></span>|<span data-ttu-id="ea709-497">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="ea709-497">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="ea709-498">エラー</span><span class="sxs-lookup"><span data-stu-id="ea709-498">Errors</span></span>

|<span data-ttu-id="ea709-499">エラー コード</span><span class="sxs-lookup"><span data-stu-id="ea709-499">Error code</span></span>|<span data-ttu-id="ea709-500">説明</span><span class="sxs-lookup"><span data-stu-id="ea709-500">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="ea709-501">要求が失敗しました。</span><span class="sxs-lookup"><span data-stu-id="ea709-501">The request has failed.</span></span> <span data-ttu-id="ea709-502">HTTP エラーコードの diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="ea709-502">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="ea709-503">Exchange サーバーがエラーを返しました。</span><span class="sxs-lookup"><span data-stu-id="ea709-503">The Exchange server returned an error.</span></span> <span data-ttu-id="ea709-504">詳細については、diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="ea709-504">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="ea709-505">ユーザーはネットワークに接続されていません。</span><span class="sxs-lookup"><span data-stu-id="ea709-505">The user is no longer connected to the network.</span></span> <span data-ttu-id="ea709-506">ネットワーク接続を確認し、やり直してください。</span><span class="sxs-lookup"><span data-stu-id="ea709-506">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ea709-507">要件</span><span class="sxs-lookup"><span data-stu-id="ea709-507">Requirements</span></span>

|<span data-ttu-id="ea709-508">要件</span><span class="sxs-lookup"><span data-stu-id="ea709-508">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="ea709-509">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ea709-509">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ea709-510">1.0</span><span class="sxs-lookup"><span data-stu-id="ea709-510">1.0</span></span> | <span data-ttu-id="ea709-511">1.3</span><span class="sxs-lookup"><span data-stu-id="ea709-511">1.3</span></span> |
|[<span data-ttu-id="ea709-512">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ea709-512">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ea709-513">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ea709-513">ReadItem</span></span> | <span data-ttu-id="ea709-514">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ea709-514">ReadItem</span></span> |
|[<span data-ttu-id="ea709-515">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ea709-515">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ea709-516">Read</span><span class="sxs-lookup"><span data-stu-id="ea709-516">Read</span></span> | <span data-ttu-id="ea709-517">Compose</span><span class="sxs-lookup"><span data-stu-id="ea709-517">Compose</span></span> |

##### <a name="example"></a><span data-ttu-id="ea709-518">例</span><span class="sxs-lookup"><span data-stu-id="ea709-518">Example</span></span>

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

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="ea709-519">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="ea709-519">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="ea709-520">ユーザーと Office アドインを識別するトークンを取得します。</span><span class="sxs-lookup"><span data-stu-id="ea709-520">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="ea709-521">`getUserIdentityTokenAsync` メソッドは、[アドインとユーザーをサード パーティのシステムで識別して認証](/outlook/add-ins/authentication)することのできるトークンを返します。</span><span class="sxs-lookup"><span data-stu-id="ea709-521">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="ea709-522">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ea709-522">Parameters</span></span>

|<span data-ttu-id="ea709-523">名前</span><span class="sxs-lookup"><span data-stu-id="ea709-523">Name</span></span>| <span data-ttu-id="ea709-524">型</span><span class="sxs-lookup"><span data-stu-id="ea709-524">Type</span></span>| <span data-ttu-id="ea709-525">属性</span><span class="sxs-lookup"><span data-stu-id="ea709-525">Attributes</span></span>| <span data-ttu-id="ea709-526">説明</span><span class="sxs-lookup"><span data-stu-id="ea709-526">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="ea709-527">function</span><span class="sxs-lookup"><span data-stu-id="ea709-527">function</span></span>||<span data-ttu-id="ea709-528">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="ea709-528">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="ea709-529">トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="ea709-529">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="ea709-530">エラーが発生した場合、 `asyncResult.error` および `asyncResult.diagnostics` のプロパティで追加情報が提供される場合があります。</span><span class="sxs-lookup"><span data-stu-id="ea709-530">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="ea709-531">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="ea709-531">Object</span></span>| <span data-ttu-id="ea709-532">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="ea709-532">&lt;optional&gt;</span></span>|<span data-ttu-id="ea709-533">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="ea709-533">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="ea709-534">エラー</span><span class="sxs-lookup"><span data-stu-id="ea709-534">Errors</span></span>

|<span data-ttu-id="ea709-535">エラー コード</span><span class="sxs-lookup"><span data-stu-id="ea709-535">Error code</span></span>|<span data-ttu-id="ea709-536">説明</span><span class="sxs-lookup"><span data-stu-id="ea709-536">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="ea709-537">要求が失敗しました。</span><span class="sxs-lookup"><span data-stu-id="ea709-537">The request has failed.</span></span> <span data-ttu-id="ea709-538">HTTP エラーコードの diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="ea709-538">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="ea709-539">Exchange サーバーがエラーを返しました。</span><span class="sxs-lookup"><span data-stu-id="ea709-539">The Exchange server returned an error.</span></span> <span data-ttu-id="ea709-540">詳細については、diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="ea709-540">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="ea709-541">ユーザーはネットワークに接続されていません。</span><span class="sxs-lookup"><span data-stu-id="ea709-541">The user is no longer connected to the network.</span></span> <span data-ttu-id="ea709-542">ネットワーク接続を確認し、やり直してください。</span><span class="sxs-lookup"><span data-stu-id="ea709-542">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ea709-543">要件</span><span class="sxs-lookup"><span data-stu-id="ea709-543">Requirements</span></span>

|<span data-ttu-id="ea709-544">要件</span><span class="sxs-lookup"><span data-stu-id="ea709-544">Requirement</span></span>| <span data-ttu-id="ea709-545">値</span><span class="sxs-lookup"><span data-stu-id="ea709-545">Value</span></span>|
|---|---|
|[<span data-ttu-id="ea709-546">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ea709-546">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ea709-547">1.0</span><span class="sxs-lookup"><span data-stu-id="ea709-547">1.0</span></span>|
|[<span data-ttu-id="ea709-548">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ea709-548">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ea709-549">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ea709-549">ReadItem</span></span>|
|[<span data-ttu-id="ea709-550">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ea709-550">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ea709-551">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ea709-551">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ea709-552">例</span><span class="sxs-lookup"><span data-stu-id="ea709-552">Example</span></span>

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

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="ea709-553">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="ea709-553">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="ea709-554">ユーザーのメールボックスをホストしている Exchange サーバー上の Exchange Web サービス (EWS) のサービスに対して非同期の要求を行います。</span><span class="sxs-lookup"><span data-stu-id="ea709-554">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="ea709-555">このメソッドは、次のシナリオではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="ea709-555">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="ea709-556">Outlook on iOS または Android の場合</span><span class="sxs-lookup"><span data-stu-id="ea709-556">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="ea709-557">アドインが Gmail のメールボックスに読み込まれる場合</span><span class="sxs-lookup"><span data-stu-id="ea709-557">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="ea709-558">このような場合は、アドインでは [REST API を使用](/outlook/add-ins/use-rest-api)して、代わりにユーザーのメールボックスにアクセスする必要があります。</span><span class="sxs-lookup"><span data-stu-id="ea709-558">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="ea709-p140">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange. See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span><span class="sxs-lookup"><span data-stu-id="ea709-p140">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange. See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="ea709-561">`makeEwsRequestAsync` メソッドでは、フォルダー関連アイテムを要求できません。</span><span class="sxs-lookup"><span data-stu-id="ea709-561">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="ea709-562">XML 要求では UTF-8 エンコードを指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="ea709-562">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="ea709-p141">`makeEwsRequestAsync` メソッドを使用するには、アドインに **ReadWriteMailbox** アクセス許可が必要です。**ReadWriteMailbox** アクセス許可と、`makeEwsRequestAsync` メソッドで呼び出せる EWS 操作の使い方については、「[ユーザーのメールボックスへのメール アドイン アクセスのアクセス許可を指定する](/outlook/add-ins/understanding-outlook-add-in-permissions)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="ea709-p141">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="ea709-565">サーバー管理者は、クライアント アクセス サーバーの EWS ディレクトリで `OAuthAuthentication` を true に設定して、`makeEwsRequestAsync` メソッドで EWS 要求を行うことができるようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="ea709-565">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="ea709-566">バージョンの相違点</span><span class="sxs-lookup"><span data-stu-id="ea709-566">Version differences</span></span>

<span data-ttu-id="ea709-567">バージョン 15.0.4535.1004 より前のバージョンの Outlook で実行しているメール アプリで `makeEwsRequestAsync` メソッドを使う場合は、エンコード値を `ISO-8859-1` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="ea709-567">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="ea709-p142">Outlook on the web でメール アプリを実行している場合は、エンコード値を設定する必要はありません。mailbox.diagnostics.hostName プロパティを使って、メール アプリを Outlook で実行しているのか、Outlook on the web で実行しているのかを確認できます。mailbox.diagnostics.hostVersion プロパティを使って、どのバージョンの Outlook を使って実行しているかを確認できます。</span><span class="sxs-lookup"><span data-stu-id="ea709-p142">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ea709-571">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ea709-571">Parameters</span></span>

|<span data-ttu-id="ea709-572">名前</span><span class="sxs-lookup"><span data-stu-id="ea709-572">Name</span></span>| <span data-ttu-id="ea709-573">型</span><span class="sxs-lookup"><span data-stu-id="ea709-573">Type</span></span>| <span data-ttu-id="ea709-574">属性</span><span class="sxs-lookup"><span data-stu-id="ea709-574">Attributes</span></span>| <span data-ttu-id="ea709-575">説明</span><span class="sxs-lookup"><span data-stu-id="ea709-575">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="ea709-576">String</span><span class="sxs-lookup"><span data-stu-id="ea709-576">String</span></span>||<span data-ttu-id="ea709-577">EWS 要求です。</span><span class="sxs-lookup"><span data-stu-id="ea709-577">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="ea709-578">function</span><span class="sxs-lookup"><span data-stu-id="ea709-578">function</span></span>||<span data-ttu-id="ea709-579">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="ea709-579">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="ea709-p143">The XML result of the EWS call is provided as a string in the `asyncResult.value` property. If the result exceeds 1 MB in size, an error message is returned instead.</span><span class="sxs-lookup"><span data-stu-id="ea709-p143">The XML result of the EWS call is provided as a string in the `asyncResult.value` property. If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="ea709-582">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="ea709-582">Object</span></span>| <span data-ttu-id="ea709-583">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="ea709-583">&lt;optional&gt;</span></span>|<span data-ttu-id="ea709-584">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="ea709-584">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ea709-585">要件</span><span class="sxs-lookup"><span data-stu-id="ea709-585">Requirements</span></span>

|<span data-ttu-id="ea709-586">要件</span><span class="sxs-lookup"><span data-stu-id="ea709-586">Requirement</span></span>| <span data-ttu-id="ea709-587">値</span><span class="sxs-lookup"><span data-stu-id="ea709-587">Value</span></span>|
|---|---|
|[<span data-ttu-id="ea709-588">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ea709-588">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ea709-589">1.0</span><span class="sxs-lookup"><span data-stu-id="ea709-589">1.0</span></span>|
|[<span data-ttu-id="ea709-590">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ea709-590">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ea709-591">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="ea709-591">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="ea709-592">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ea709-592">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ea709-593">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ea709-593">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ea709-594">例</span><span class="sxs-lookup"><span data-stu-id="ea709-594">Example</span></span>

<span data-ttu-id="ea709-595">次の例は、`GetItem` 操作を使ってアイテムの件名を取得するため、`makeEwsRequestAsync` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="ea709-595">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="ea709-596">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="ea709-596">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="ea709-597">サポートされているイベントの種類のイベント ハンドラーを削除します。</span><span class="sxs-lookup"><span data-stu-id="ea709-597">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="ea709-598">現在、サポートされているイベントの種類は `Office.EventType.ItemChanged` だけです。</span><span class="sxs-lookup"><span data-stu-id="ea709-598">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ea709-599">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ea709-599">Parameters</span></span>

| <span data-ttu-id="ea709-600">名前</span><span class="sxs-lookup"><span data-stu-id="ea709-600">Name</span></span> | <span data-ttu-id="ea709-601">型</span><span class="sxs-lookup"><span data-stu-id="ea709-601">Type</span></span> | <span data-ttu-id="ea709-602">属性</span><span class="sxs-lookup"><span data-stu-id="ea709-602">Attributes</span></span> | <span data-ttu-id="ea709-603">説明</span><span class="sxs-lookup"><span data-stu-id="ea709-603">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="ea709-604">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="ea709-604">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="ea709-605">ハンドラーを取り消すイベント。</span><span class="sxs-lookup"><span data-stu-id="ea709-605">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="ea709-606">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="ea709-606">Object</span></span> | <span data-ttu-id="ea709-607">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ea709-607">&lt;optional&gt;</span></span> | <span data-ttu-id="ea709-608">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="ea709-608">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="ea709-609">Object</span><span class="sxs-lookup"><span data-stu-id="ea709-609">Object</span></span> | <span data-ttu-id="ea709-610">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ea709-610">&lt;optional&gt;</span></span> | <span data-ttu-id="ea709-611">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="ea709-611">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="ea709-612">function</span><span class="sxs-lookup"><span data-stu-id="ea709-612">function</span></span>| <span data-ttu-id="ea709-613">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ea709-613">&lt;optional&gt;</span></span>|<span data-ttu-id="ea709-614">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="ea709-614">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ea709-615">要件</span><span class="sxs-lookup"><span data-stu-id="ea709-615">Requirements</span></span>

|<span data-ttu-id="ea709-616">要件</span><span class="sxs-lookup"><span data-stu-id="ea709-616">Requirement</span></span>| <span data-ttu-id="ea709-617">値</span><span class="sxs-lookup"><span data-stu-id="ea709-617">Value</span></span>|
|---|---|
|[<span data-ttu-id="ea709-618">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ea709-618">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ea709-619">1.5</span><span class="sxs-lookup"><span data-stu-id="ea709-619">1.5</span></span> |
|[<span data-ttu-id="ea709-620">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ea709-620">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ea709-621">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ea709-621">ReadItem</span></span> |
|[<span data-ttu-id="ea709-622">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ea709-622">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ea709-623">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ea709-623">Compose or Read</span></span>|
