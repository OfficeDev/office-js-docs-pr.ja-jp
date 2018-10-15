
# <a name="mailbox"></a><span data-ttu-id="8bad4-101">mailbox</span><span class="sxs-lookup"><span data-stu-id="8bad4-101">mailbox</span></span>

### <span data-ttu-id="8bad4-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span><span class="sxs-lookup"><span data-stu-id="8bad4-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span></span>

<span data-ttu-id="8bad4-104">Microsoft Outlook と Microsoft Outlook on the web の Outlook アドイン オブジェクト モデルへのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="8bad4-104">Provides access to the Outlook Add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8bad4-105">要件</span><span class="sxs-lookup"><span data-stu-id="8bad4-105">Requirements</span></span>

|<span data-ttu-id="8bad4-106">要件</span><span class="sxs-lookup"><span data-stu-id="8bad4-106">Requirement</span></span>| <span data-ttu-id="8bad4-107">値</span><span class="sxs-lookup"><span data-stu-id="8bad4-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="8bad4-108">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="8bad4-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8bad4-109">1.0</span><span class="sxs-lookup"><span data-stu-id="8bad4-109">1.0</span></span>|
|[<span data-ttu-id="8bad4-110">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="8bad4-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8bad4-111">制限あり</span><span class="sxs-lookup"><span data-stu-id="8bad4-111">Restricted</span></span>|
|[<span data-ttu-id="8bad4-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8bad4-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8bad4-113">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="8bad4-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="8bad4-114">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="8bad4-114">Members and methods</span></span>

| <span data-ttu-id="8bad4-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="8bad4-115">Member</span></span> | <span data-ttu-id="8bad4-116">種類</span><span class="sxs-lookup"><span data-stu-id="8bad4-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="8bad4-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="8bad4-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="8bad4-118">メンバー</span><span class="sxs-lookup"><span data-stu-id="8bad4-118">Member</span></span> |
| [<span data-ttu-id="8bad4-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="8bad4-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="8bad4-120">メンバー</span><span class="sxs-lookup"><span data-stu-id="8bad4-120">Member</span></span> |
| [<span data-ttu-id="8bad4-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="8bad4-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="8bad4-122">メソッド</span><span class="sxs-lookup"><span data-stu-id="8bad4-122">Method</span></span> |
| [<span data-ttu-id="8bad4-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="8bad4-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="8bad4-124">メソッド</span><span class="sxs-lookup"><span data-stu-id="8bad4-124">Method</span></span> |
| [<span data-ttu-id="8bad4-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="8bad4-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) | <span data-ttu-id="8bad4-126">メソッド</span><span class="sxs-lookup"><span data-stu-id="8bad4-126">Method</span></span> |
| [<span data-ttu-id="8bad4-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="8bad4-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="8bad4-128">メソッド</span><span class="sxs-lookup"><span data-stu-id="8bad4-128">Method</span></span> |
| [<span data-ttu-id="8bad4-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="8bad4-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="8bad4-130">メソッド</span><span class="sxs-lookup"><span data-stu-id="8bad4-130">Method</span></span> |
| [<span data-ttu-id="8bad4-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="8bad4-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="8bad4-132">メソッド</span><span class="sxs-lookup"><span data-stu-id="8bad4-132">Method</span></span> |
| [<span data-ttu-id="8bad4-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="8bad4-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="8bad4-134">メソッド</span><span class="sxs-lookup"><span data-stu-id="8bad4-134">Method</span></span> |
| [<span data-ttu-id="8bad4-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="8bad4-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="8bad4-136">メソッド</span><span class="sxs-lookup"><span data-stu-id="8bad4-136">Method</span></span> |
| [<span data-ttu-id="8bad4-137">displayNewMessageForm</span><span class="sxs-lookup"><span data-stu-id="8bad4-137">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="8bad4-138">メソッド</span><span class="sxs-lookup"><span data-stu-id="8bad4-138">Method</span></span> |
| [<span data-ttu-id="8bad4-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="8bad4-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="8bad4-140">メソッド</span><span class="sxs-lookup"><span data-stu-id="8bad4-140">Method</span></span> |
| [<span data-ttu-id="8bad4-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="8bad4-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="8bad4-142">メソッド</span><span class="sxs-lookup"><span data-stu-id="8bad4-142">Method</span></span> |
| [<span data-ttu-id="8bad4-143">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="8bad4-143">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="8bad4-144">メソッド</span><span class="sxs-lookup"><span data-stu-id="8bad4-144">Method</span></span> |
| [<span data-ttu-id="8bad4-145">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="8bad4-145">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="8bad4-146">メソッド</span><span class="sxs-lookup"><span data-stu-id="8bad4-146">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="8bad4-147">名前空間</span><span class="sxs-lookup"><span data-stu-id="8bad4-147">Namespaces</span></span>

<span data-ttu-id="8bad4-148">[diagnostics](Office.context.mailbox.diagnostics.md): Outlook アドインに診断情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="8bad4-148">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="8bad4-149">[item](Office.context.mailbox.item.md):Outlook アドインのメッセージや予定にアクセスするためのメソッドとプロパティを提供します。</span><span class="sxs-lookup"><span data-stu-id="8bad4-149">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="8bad4-150">[userProfile](Office.context.mailbox.userProfile.md):Outlook アドインのユーザーに関する情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="8bad4-150">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="8bad4-151">メンバー</span><span class="sxs-lookup"><span data-stu-id="8bad4-151">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="8bad4-152">ewsUrl: 文字列</span><span class="sxs-lookup"><span data-stu-id="8bad4-152">ewsUrl :String</span></span>

<span data-ttu-id="8bad4-p102">このメール アカウントの Exchange Web サービス (EWS) エンドポイントの URL を取得します。閲覧モードのみです。</span><span class="sxs-lookup"><span data-stu-id="8bad4-p102">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="8bad4-155">このメンバーは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="8bad4-155">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="8bad4-p103">`ewsUrl` 値は、リモート サービスで、ユーザーのメールボックスに EWS 呼び出しを行うために使用することができます。たとえば、[選択した項目から添付ファイルを取得する](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="8bad4-p103">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="8bad4-158">閲覧モードで `ewsUrl` メンバーを呼び出すには、アプリでマニフェスト内に **ReadItem** アクセス許可を指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="8bad4-158">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="8bad4-p104">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback)メソッドを呼び出してから、`ewsUrl`メンバーを使用する必要があります。アプリには、`saveAsync`メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="8bad4-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="8bad4-161">型:</span><span class="sxs-lookup"><span data-stu-id="8bad4-161">Type:</span></span>

*   <span data-ttu-id="8bad4-162">文字列</span><span class="sxs-lookup"><span data-stu-id="8bad4-162">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8bad4-163">要件</span><span class="sxs-lookup"><span data-stu-id="8bad4-163">Requirements</span></span>

|<span data-ttu-id="8bad4-164">要件</span><span class="sxs-lookup"><span data-stu-id="8bad4-164">Requirement</span></span>| <span data-ttu-id="8bad4-165">値</span><span class="sxs-lookup"><span data-stu-id="8bad4-165">Value</span></span>|
|---|---|
|[<span data-ttu-id="8bad4-166">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="8bad4-166">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8bad4-167">1.0</span><span class="sxs-lookup"><span data-stu-id="8bad4-167">1.0</span></span>|
|[<span data-ttu-id="8bad4-168">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="8bad4-168">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8bad4-169">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8bad4-169">ReadItem</span></span>|
|[<span data-ttu-id="8bad4-170">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8bad4-170">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8bad4-171">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="8bad4-171">Compose or read</span></span>|

#### <a name="resturl-string"></a><span data-ttu-id="8bad4-172">restUrl: 文字列</span><span class="sxs-lookup"><span data-stu-id="8bad4-172">restUrl :String</span></span>

<span data-ttu-id="8bad4-173">この電子メール アカウントの REST エンドポイントの URL を取得します。</span><span class="sxs-lookup"><span data-stu-id="8bad4-173">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="8bad4-174">`restUrl` 値は、ユーザーのメールボックスに [REST API](https://docs.microsoft.com/outlook/rest/) 呼び出しを行うために使用することができます。</span><span class="sxs-lookup"><span data-stu-id="8bad4-174">The `restUrl` value can be used to make [REST API](https://docs.microsoft.com/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="8bad4-175">閲覧モードで `restUrl` メンバーを呼び出すには、アプリでマニフェスト内に **ReadItem** アクセス許可を指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="8bad4-175">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="8bad4-p105">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback)メソッドを呼び出してから、`restUrl`メンバーを使用する必要があります。アプリには、`saveAsync`メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="8bad4-p105">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="8bad4-178">型:</span><span class="sxs-lookup"><span data-stu-id="8bad4-178">Type:</span></span>

*   <span data-ttu-id="8bad4-179">文字列</span><span class="sxs-lookup"><span data-stu-id="8bad4-179">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8bad4-180">要件</span><span class="sxs-lookup"><span data-stu-id="8bad4-180">Requirements</span></span>

|<span data-ttu-id="8bad4-181">要件</span><span class="sxs-lookup"><span data-stu-id="8bad4-181">Requirement</span></span>| <span data-ttu-id="8bad4-182">値</span><span class="sxs-lookup"><span data-stu-id="8bad4-182">Value</span></span>|
|---|---|
|[<span data-ttu-id="8bad4-183">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="8bad4-183">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8bad4-184">1.5</span><span class="sxs-lookup"><span data-stu-id="8bad4-184">1.5</span></span> |
|[<span data-ttu-id="8bad4-185">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="8bad4-185">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8bad4-186">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8bad4-186">ReadItem</span></span>|
|[<span data-ttu-id="8bad4-187">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8bad4-187">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8bad4-188">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="8bad4-188">Compose or read</span></span>|

### <a name="methods"></a><span data-ttu-id="8bad4-189">メソッド</span><span class="sxs-lookup"><span data-stu-id="8bad4-189">Methods</span></span>

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="8bad4-190">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="8bad4-190">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="8bad4-191">サポートされているイベントのイベント ハンドラを追加します。</span><span class="sxs-lookup"><span data-stu-id="8bad4-191">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="8bad4-192">現在、サポートされているイベントの種類は、`Office.EventType.ItemChanged`と`Office.EventType.OfficeThemeChanged`。</span><span class="sxs-lookup"><span data-stu-id="8bad4-192">Currently, the supported event types are `Office.EventType.ItemChanged` and `Office.EventType.OfficeThemeChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8bad4-193">パラメータ:</span><span class="sxs-lookup"><span data-stu-id="8bad4-193">Parameters:</span></span>

| <span data-ttu-id="8bad4-194">名前</span><span class="sxs-lookup"><span data-stu-id="8bad4-194">Name</span></span> | <span data-ttu-id="8bad4-195">種類</span><span class="sxs-lookup"><span data-stu-id="8bad4-195">Type</span></span> | <span data-ttu-id="8bad4-196">属性</span><span class="sxs-lookup"><span data-stu-id="8bad4-196">Attributes</span></span> | <span data-ttu-id="8bad4-197">説明</span><span class="sxs-lookup"><span data-stu-id="8bad4-197">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="8bad4-198">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="8bad4-198">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="8bad4-199">ハンドラを呼び出す必要のあるイベント。</span><span class="sxs-lookup"><span data-stu-id="8bad4-199">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="8bad4-200">関数</span><span class="sxs-lookup"><span data-stu-id="8bad4-200">Function</span></span> || <span data-ttu-id="8bad4-p106">イベントを処理する関数。この関数は、オブジェクト リテラルである単一パラメータを受け入れる必要があります。パラメータの `type` プロパティは、`addHandlerAsync` に渡される `eventType` パラメータと一致します。</span><span class="sxs-lookup"><span data-stu-id="8bad4-p106">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="8bad4-204">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="8bad4-204">Object</span></span> | <span data-ttu-id="8bad4-205">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="8bad4-205">&lt;optional&gt;</span></span> | <span data-ttu-id="8bad4-206">以下のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="8bad4-206">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="8bad4-207">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="8bad4-207">Object</span></span> | <span data-ttu-id="8bad4-208">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="8bad4-208">&lt;optional&gt;</span></span> | <span data-ttu-id="8bad4-209">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="8bad4-209">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="8bad4-210">関数</span><span class="sxs-lookup"><span data-stu-id="8bad4-210">function</span></span>| <span data-ttu-id="8bad4-211">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="8bad4-211">&lt;optional&gt;</span></span>|<span data-ttu-id="8bad4-212">メソッドが完了すると、`callback` パラメータに渡された関数が、単一パラメータ `asyncResult` で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="8bad4-212">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8bad4-213">要件</span><span class="sxs-lookup"><span data-stu-id="8bad4-213">Requirements</span></span>

|<span data-ttu-id="8bad4-214">要件</span><span class="sxs-lookup"><span data-stu-id="8bad4-214">Requirement</span></span>| <span data-ttu-id="8bad4-215">値</span><span class="sxs-lookup"><span data-stu-id="8bad4-215">Value</span></span>|
|---|---|
|[<span data-ttu-id="8bad4-216">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="8bad4-216">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8bad4-217">1.5</span><span class="sxs-lookup"><span data-stu-id="8bad4-217">1.5</span></span> |
|[<span data-ttu-id="8bad4-218">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="8bad4-218">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8bad4-219">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8bad4-219">ReadItem</span></span> |
|[<span data-ttu-id="8bad4-220">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8bad4-220">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8bad4-221">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="8bad4-221">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8bad4-222">例</span><span class="sxs-lookup"><span data-stu-id="8bad4-222">Example</span></span>

```
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, loadNewItem, function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        // Handle error
      }
    });
  });
};

function loadNewItem(eventArgs) {
  // Load the properties of the newly selected item
  loadProps(Office.context.mailbox.item);
};
```

####  <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="8bad4-223">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="8bad4-223">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="8bad4-224">REST 用に書式設定された項目 ID を EWS 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="8bad4-224">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="8bad4-225">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="8bad4-225">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="8bad4-p107">REST API ([Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) や [Microsoft Graph](http://graph.microsoft.io/) など) 経由で取得された項目 ID は、Exchange Web サービス (EWS) で使用される形式とは異なる形式を使用します。`convertToEwsId` メソッドは、REST 形式の ID を EWS 用の適切な形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="8bad4-p107">Item IDs retrieved via a REST API (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](http://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8bad4-228">パラメータ:</span><span class="sxs-lookup"><span data-stu-id="8bad4-228">Parameters:</span></span>

|<span data-ttu-id="8bad4-229">名前</span><span class="sxs-lookup"><span data-stu-id="8bad4-229">Name</span></span>| <span data-ttu-id="8bad4-230">種類</span><span class="sxs-lookup"><span data-stu-id="8bad4-230">Type</span></span>| <span data-ttu-id="8bad4-231">説明</span><span class="sxs-lookup"><span data-stu-id="8bad4-231">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="8bad4-232">文字列</span><span class="sxs-lookup"><span data-stu-id="8bad4-232">String</span></span>|<span data-ttu-id="8bad4-233">Outlook REST API 用に書式設定された項目 ID</span><span class="sxs-lookup"><span data-stu-id="8bad4-233">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="8bad4-234">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="8bad4-234">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.restversion)|<span data-ttu-id="8bad4-235">項目 ID の取得に使用された Outlook REST API のバージョンを示す値。</span><span class="sxs-lookup"><span data-stu-id="8bad4-235">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8bad4-236">要件</span><span class="sxs-lookup"><span data-stu-id="8bad4-236">Requirements</span></span>

|<span data-ttu-id="8bad4-237">要件</span><span class="sxs-lookup"><span data-stu-id="8bad4-237">Requirement</span></span>| <span data-ttu-id="8bad4-238">値</span><span class="sxs-lookup"><span data-stu-id="8bad4-238">Value</span></span>|
|---|---|
|[<span data-ttu-id="8bad4-239">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="8bad4-239">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8bad4-240">1.3</span><span class="sxs-lookup"><span data-stu-id="8bad4-240">1.3</span></span>|
|[<span data-ttu-id="8bad4-241">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="8bad4-241">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8bad4-242">制限あり</span><span class="sxs-lookup"><span data-stu-id="8bad4-242">Restricted</span></span>|
|[<span data-ttu-id="8bad4-243">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8bad4-243">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8bad4-244">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="8bad4-244">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8bad4-245">戻り値 :</span><span class="sxs-lookup"><span data-stu-id="8bad4-245">Returns:</span></span>

<span data-ttu-id="8bad4-246">種類: 文字列</span><span class="sxs-lookup"><span data-stu-id="8bad4-246">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="8bad4-247">例</span><span class="sxs-lookup"><span data-stu-id="8bad4-247">Example</span></span>

```
// Get an item's ID from a REST API
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the
// Outlook Mail API
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime"></a><span data-ttu-id="8bad4-248">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_7/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="8bad4-248">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_7/office.LocalClientTime)}</span></span>

<span data-ttu-id="8bad4-249">クライアントの現地時間の時間情報が含まれている辞書を取得します。</span><span class="sxs-lookup"><span data-stu-id="8bad4-249">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="8bad4-p108">Outlook 用メール アプリや Outlook Web App で使う日付と時刻は、異なるタイム ゾーンを使うことができます。Outlook では、クライアント コンピューターのタイム ゾーンを使います。Outlook Web App では、Exchange 管理センター (EAC) で設定されたタイム ゾーンを使います。ユーザー インターフェイスに表示される値が、常にユーザーが期待するタイム ゾーンと一致するように日付と時刻の値を処理する必要があります。</span><span class="sxs-lookup"><span data-stu-id="8bad4-p108">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="8bad4-p109">Outlook でメール アプリが実行されている場合、`convertToLocalClientTime` メソッドは、クライアント コンピューターのタイム ゾーンに設定された値で辞書オブジェクトを返します。Outlook Web Apps でメール アプリが実行されている場合、`convertToLocalClientTime` メソッドは、EAC で指定されたタイム ゾーンに設定された値で辞書オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="8bad4-p109">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8bad4-255">パラメータ:</span><span class="sxs-lookup"><span data-stu-id="8bad4-255">Parameters:</span></span>

|<span data-ttu-id="8bad4-256">名前</span><span class="sxs-lookup"><span data-stu-id="8bad4-256">Name</span></span>| <span data-ttu-id="8bad4-257">種類</span><span class="sxs-lookup"><span data-stu-id="8bad4-257">Type</span></span>| <span data-ttu-id="8bad4-258">説明</span><span class="sxs-lookup"><span data-stu-id="8bad4-258">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="8bad4-259">Date</span><span class="sxs-lookup"><span data-stu-id="8bad4-259">Date</span></span>|<span data-ttu-id="8bad4-260">Date オブジェクト</span><span class="sxs-lookup"><span data-stu-id="8bad4-260">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8bad4-261">要件</span><span class="sxs-lookup"><span data-stu-id="8bad4-261">Requirements</span></span>

|<span data-ttu-id="8bad4-262">要件</span><span class="sxs-lookup"><span data-stu-id="8bad4-262">Requirement</span></span>| <span data-ttu-id="8bad4-263">値</span><span class="sxs-lookup"><span data-stu-id="8bad4-263">Value</span></span>|
|---|---|
|[<span data-ttu-id="8bad4-264">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="8bad4-264">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8bad4-265">1.0</span><span class="sxs-lookup"><span data-stu-id="8bad4-265">1.0</span></span>|
|[<span data-ttu-id="8bad4-266">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="8bad4-266">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8bad4-267">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8bad4-267">ReadItem</span></span>|
|[<span data-ttu-id="8bad4-268">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8bad4-268">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8bad4-269">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="8bad4-269">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8bad4-270">戻り値 :</span><span class="sxs-lookup"><span data-stu-id="8bad4-270">Returns:</span></span>

<span data-ttu-id="8bad4-271">種類:[LocalClientTime](/javascript/api/outlook_1_7/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="8bad4-271">Type: [LocalClientTime](/javascript/api/outlook_1_7/office.LocalClientTime)</span></span>

####  <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="8bad4-272">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="8bad4-272">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="8bad4-273">EWS 用に書式設定された項目 ID を REST 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="8bad4-273">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="8bad4-274">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="8bad4-274">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="8bad4-p110">EWS 経由または `itemId` プロパティ経由で取得される項目 ID では、REST API ([Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) または [Microsoft Graph](http://graph.microsoft.io/) など) で使用される形式とは異なる形式を使用します。`convertToRestId` メソッドは、EWS 形式の ID を REST 用の適切な形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="8bad4-p110">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](http://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8bad4-277">パラメータ:</span><span class="sxs-lookup"><span data-stu-id="8bad4-277">Parameters:</span></span>

|<span data-ttu-id="8bad4-278">名前</span><span class="sxs-lookup"><span data-stu-id="8bad4-278">Name</span></span>| <span data-ttu-id="8bad4-279">種類</span><span class="sxs-lookup"><span data-stu-id="8bad4-279">Type</span></span>| <span data-ttu-id="8bad4-280">説明</span><span class="sxs-lookup"><span data-stu-id="8bad4-280">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="8bad4-281">文字列</span><span class="sxs-lookup"><span data-stu-id="8bad4-281">String</span></span>|<span data-ttu-id="8bad4-282">Exchange Web サービス (EWS) 用に書式設定された項目 ID</span><span class="sxs-lookup"><span data-stu-id="8bad4-282">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="8bad4-283">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="8bad4-283">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.restversion)|<span data-ttu-id="8bad4-284">変換後の ID とともに使用される Outlook REST API のバージョンを示す値。</span><span class="sxs-lookup"><span data-stu-id="8bad4-284">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8bad4-285">要件</span><span class="sxs-lookup"><span data-stu-id="8bad4-285">Requirements</span></span>

|<span data-ttu-id="8bad4-286">要件</span><span class="sxs-lookup"><span data-stu-id="8bad4-286">Requirement</span></span>| <span data-ttu-id="8bad4-287">値</span><span class="sxs-lookup"><span data-stu-id="8bad4-287">Value</span></span>|
|---|---|
|[<span data-ttu-id="8bad4-288">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="8bad4-288">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8bad4-289">1.3</span><span class="sxs-lookup"><span data-stu-id="8bad4-289">1.3</span></span>|
|[<span data-ttu-id="8bad4-290">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="8bad4-290">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8bad4-291">制限あり</span><span class="sxs-lookup"><span data-stu-id="8bad4-291">Restricted</span></span>|
|[<span data-ttu-id="8bad4-292">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8bad4-292">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8bad4-293">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="8bad4-293">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8bad4-294">戻り値 :</span><span class="sxs-lookup"><span data-stu-id="8bad4-294">Returns:</span></span>

<span data-ttu-id="8bad4-295">種類: 文字列</span><span class="sxs-lookup"><span data-stu-id="8bad4-295">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="8bad4-296">例</span><span class="sxs-lookup"><span data-stu-id="8bad4-296">Example</span></span>

```
// Get the currently selected item's ID
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the
// Outlook Mail API
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="8bad4-297">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="8bad4-297">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="8bad4-298">時間情報が含まれている辞書から Date オブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="8bad4-298">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="8bad4-299">`convertToUtcClientTime` メソッドは、ローカルの日付と時刻が含まれる辞書を、ローカルの日付と時刻の正しい値をもつ Date オブジェクトに変換します。</span><span class="sxs-lookup"><span data-stu-id="8bad4-299">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8bad4-300">パラメータ:</span><span class="sxs-lookup"><span data-stu-id="8bad4-300">Parameters:</span></span>

|<span data-ttu-id="8bad4-301">名前</span><span class="sxs-lookup"><span data-stu-id="8bad4-301">Name</span></span>| <span data-ttu-id="8bad4-302">種類</span><span class="sxs-lookup"><span data-stu-id="8bad4-302">Type</span></span>| <span data-ttu-id="8bad4-303">説明</span><span class="sxs-lookup"><span data-stu-id="8bad4-303">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="8bad4-304">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="8bad4-304">LocalClientTime</span></span>](/javascript/api/outlook_1_7/office.LocalClientTime)|<span data-ttu-id="8bad4-305">変換するローカル時刻の値。</span><span class="sxs-lookup"><span data-stu-id="8bad4-305">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8bad4-306">要件</span><span class="sxs-lookup"><span data-stu-id="8bad4-306">Requirements</span></span>

|<span data-ttu-id="8bad4-307">要件</span><span class="sxs-lookup"><span data-stu-id="8bad4-307">Requirement</span></span>| <span data-ttu-id="8bad4-308">値</span><span class="sxs-lookup"><span data-stu-id="8bad4-308">Value</span></span>|
|---|---|
|[<span data-ttu-id="8bad4-309">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="8bad4-309">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8bad4-310">1.0</span><span class="sxs-lookup"><span data-stu-id="8bad4-310">1.0</span></span>|
|[<span data-ttu-id="8bad4-311">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="8bad4-311">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8bad4-312">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8bad4-312">ReadItem</span></span>|
|[<span data-ttu-id="8bad4-313">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8bad4-313">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8bad4-314">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="8bad4-314">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8bad4-315">戻り値 :</span><span class="sxs-lookup"><span data-stu-id="8bad4-315">Returns:</span></span>

<span data-ttu-id="8bad4-316">時間が UTC で表現された Date オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="8bad4-316">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="8bad4-317">

<dt>種類</dt>

</span><span class="sxs-lookup"><span data-stu-id="8bad4-317">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="8bad4-318">Date</span><span class="sxs-lookup"><span data-stu-id="8bad4-318">Date</span></span></dd>

</dl>

####  <a name="displayappointmentformitemid"></a><span data-ttu-id="8bad4-319">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="8bad4-319">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="8bad4-320">既存の予定表の予定を表示します。</span><span class="sxs-lookup"><span data-stu-id="8bad4-320">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="8bad4-321">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="8bad4-321">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="8bad4-322">`displayAppointmentForm` メソッドは、デスクトップ上の新しいウィンドウまたはモバイル デバイス上のダイアログ ボックスで、既存の予定表の予定を開きます。</span><span class="sxs-lookup"><span data-stu-id="8bad4-322">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="8bad4-p111">Outlook for Mac では、このメソッドを使用して、定期的に繰り返される予定の一部ではない単発の予定、または定期的に繰り替えされる予定の元の予定を表示できます。ただし、一連の予定のインスタンスは表示できません。これは、Outlook for Mac では、定期的に繰り返されるインスタンスのプロパティ  (項目 ID を含む) にアクセスできないためです。</span><span class="sxs-lookup"><span data-stu-id="8bad4-p111">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="8bad4-325">Outlook Web App では、このメソッドで、指定されたフォームの本文が 32 KB 以下の文字数の場合に、そのフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="8bad4-325">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="8bad4-326">指定した項目識別子が既存の予定を識別しない場合には、クライアント コンピュータまたはデバイスで空白のウィンドウが開き、エラー メッセージは返されません。</span><span class="sxs-lookup"><span data-stu-id="8bad4-326">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8bad4-327">パラメータ:</span><span class="sxs-lookup"><span data-stu-id="8bad4-327">Parameters:</span></span>

|<span data-ttu-id="8bad4-328">名前</span><span class="sxs-lookup"><span data-stu-id="8bad4-328">Name</span></span>| <span data-ttu-id="8bad4-329">種類</span><span class="sxs-lookup"><span data-stu-id="8bad4-329">Type</span></span>| <span data-ttu-id="8bad4-330">説明</span><span class="sxs-lookup"><span data-stu-id="8bad4-330">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="8bad4-331">文字列</span><span class="sxs-lookup"><span data-stu-id="8bad4-331">String</span></span>|<span data-ttu-id="8bad4-332">既存の予定表の予定の Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="8bad4-332">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8bad4-333">要件</span><span class="sxs-lookup"><span data-stu-id="8bad4-333">Requirements</span></span>

|<span data-ttu-id="8bad4-334">要件</span><span class="sxs-lookup"><span data-stu-id="8bad4-334">Requirement</span></span>| <span data-ttu-id="8bad4-335">値</span><span class="sxs-lookup"><span data-stu-id="8bad4-335">Value</span></span>|
|---|---|
|[<span data-ttu-id="8bad4-336">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="8bad4-336">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8bad4-337">1.0</span><span class="sxs-lookup"><span data-stu-id="8bad4-337">1.0</span></span>|
|[<span data-ttu-id="8bad4-338">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="8bad4-338">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8bad4-339">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8bad4-339">ReadItem</span></span>|
|[<span data-ttu-id="8bad4-340">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8bad4-340">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8bad4-341">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="8bad4-341">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8bad4-342">例</span><span class="sxs-lookup"><span data-stu-id="8bad4-342">Example</span></span>

```
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  <a name="displaymessageformitemid"></a><span data-ttu-id="8bad4-343">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="8bad4-343">displayMessageForm(itemId)</span></span>

<span data-ttu-id="8bad4-344">既存のメッセージを表示します。</span><span class="sxs-lookup"><span data-stu-id="8bad4-344">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="8bad4-345">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="8bad4-345">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="8bad4-346">`displayMessageForm` メソッドは、デスクトップ上の新しいウィンドウまたはモバイル デバイス上のダイアログ ボックスで既存のメッセージを開きます。</span><span class="sxs-lookup"><span data-stu-id="8bad4-346">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="8bad4-347">Outlook Web App では、このメソッドは、指定されたフォームの本文が 32 KB 以下の文字数の場合に、そのフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="8bad4-347">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="8bad4-348">指定した項目識別子が既存のメッセージを識別しない場合には、クラアント コンピュータでメッセージは表示されず、エラー メッセージは返されません。</span><span class="sxs-lookup"><span data-stu-id="8bad4-348">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="8bad4-p112">予定を表す `itemId` が含まれる `displayMessageForm` を使用しないでください。`displayAppointmentForm` メソッドを使用して既存の予定を表示し、`displayNewAppointmentForm` を使用して新しい予定を作成するフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="8bad4-p112">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8bad4-351">パラメータ:</span><span class="sxs-lookup"><span data-stu-id="8bad4-351">Parameters:</span></span>

|<span data-ttu-id="8bad4-352">名前</span><span class="sxs-lookup"><span data-stu-id="8bad4-352">Name</span></span>| <span data-ttu-id="8bad4-353">種類</span><span class="sxs-lookup"><span data-stu-id="8bad4-353">Type</span></span>| <span data-ttu-id="8bad4-354">説明</span><span class="sxs-lookup"><span data-stu-id="8bad4-354">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="8bad4-355">文字列</span><span class="sxs-lookup"><span data-stu-id="8bad4-355">String</span></span>|<span data-ttu-id="8bad4-356">既存のメッセージの Exchange Web サービス(EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="8bad4-356">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8bad4-357">要件</span><span class="sxs-lookup"><span data-stu-id="8bad4-357">Requirements</span></span>

|<span data-ttu-id="8bad4-358">要件</span><span class="sxs-lookup"><span data-stu-id="8bad4-358">Requirement</span></span>| <span data-ttu-id="8bad4-359">値</span><span class="sxs-lookup"><span data-stu-id="8bad4-359">Value</span></span>|
|---|---|
|[<span data-ttu-id="8bad4-360">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="8bad4-360">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8bad4-361">1.0</span><span class="sxs-lookup"><span data-stu-id="8bad4-361">1.0</span></span>|
|[<span data-ttu-id="8bad4-362">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="8bad4-362">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8bad4-363">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8bad4-363">ReadItem</span></span>|
|[<span data-ttu-id="8bad4-364">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8bad4-364">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8bad4-365">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="8bad4-365">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8bad4-366">例</span><span class="sxs-lookup"><span data-stu-id="8bad4-366">Example</span></span>

```
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="8bad4-367">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="8bad4-367">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="8bad4-368">新しい予定表の予定を作成するためのフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="8bad4-368">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="8bad4-369">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="8bad4-369">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="8bad4-p113">`displayNewAppointmentForm` メソッドを使用すると、ユーザーが新しい予定または会議を作成できるフォームが開きます。パラメータを指定すると、予定のフォーム フィールドにパラメータの内容が自動的に入力されます。</span><span class="sxs-lookup"><span data-stu-id="8bad4-p113">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="8bad4-p114">Outlook Web App および OWA for Devices では、このメソッドは出席者フィールドが含まれるフォームを常に表示します。入力因数として出席者を指定しない場合には、このメソッドは [**保存**] ボタンのあるフォームを表示します。出席者を指定した場合には、フォームにはその出席者と [**送信**] ボタンが含まれます。</span><span class="sxs-lookup"><span data-stu-id="8bad4-p114">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="8bad4-p115">Outlook リッチ クライアントおよび Outlook RT では、`requiredAttendees`、`optionalAttendees` または `resources` パラメータに出席者またはリソースを指定した場合、このメソッドは [**送信**] ボタンがある会議フォームを表示します。受信者を指定しない場合には、このメソッドは [**保存して閉じる**] ボタンがある予定フォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="8bad4-p115">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="8bad4-377">パラメータのいずれかが指定されたサイズ制限を超えた場合、または不明なパラメータ名が指定された場合には、例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="8bad4-377">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8bad4-378">パラメータ:</span><span class="sxs-lookup"><span data-stu-id="8bad4-378">Parameters:</span></span>

> [!NOTE]
> <span data-ttu-id="8bad4-379">すべてのパラメータは省略可能です。</span><span class="sxs-lookup"><span data-stu-id="8bad4-379">Note: All parameters are optional.</span></span>

|<span data-ttu-id="8bad4-380">名前</span><span class="sxs-lookup"><span data-stu-id="8bad4-380">Name</span></span>| <span data-ttu-id="8bad4-381">種類</span><span class="sxs-lookup"><span data-stu-id="8bad4-381">Type</span></span>| <span data-ttu-id="8bad4-382">説明</span><span class="sxs-lookup"><span data-stu-id="8bad4-382">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="8bad4-383">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="8bad4-383">Object</span></span> | <span data-ttu-id="8bad4-384">新しい予定を記述するパラメータの辞書。</span><span class="sxs-lookup"><span data-stu-id="8bad4-384">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="8bad4-385">配列。&lt;文字列&gt; | 配列。&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="8bad4-385">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="8bad4-p116">予定への各必須出席者の電子メール アドレスを含む文字列の配列、または出席者の `EmailAddressDetails` オブジェクトを含む配列。この配列は、最大 100 エントリまでに制限されています。</span><span class="sxs-lookup"><span data-stu-id="8bad4-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="8bad4-388">配列。&lt;文字列&gt; | 配列。&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="8bad4-388">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="8bad4-p117">予定への各任意出席者の電子メール アドレスを含む文字列の配列、または出席者の `EmailAddressDetails` オブジェクトを含む配列。この配列は、最大 100 エントリまでに制限されています。</span><span class="sxs-lookup"><span data-stu-id="8bad4-p117">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="8bad4-391">Date</span><span class="sxs-lookup"><span data-stu-id="8bad4-391">Date</span></span> | <span data-ttu-id="8bad4-392">予定の開始日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="8bad4-392">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="8bad4-393">Date</span><span class="sxs-lookup"><span data-stu-id="8bad4-393">Date</span></span> | <span data-ttu-id="8bad4-394">予定の終了日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="8bad4-394">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="8bad4-395">文字列</span><span class="sxs-lookup"><span data-stu-id="8bad4-395">String</span></span> | <span data-ttu-id="8bad4-p118">予定の場所を含む文字列。文字列は最大 255 文字までに制限されています。</span><span class="sxs-lookup"><span data-stu-id="8bad4-p118">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="8bad4-398">配列。&lt; 文字列&gt;</span><span class="sxs-lookup"><span data-stu-id="8bad4-398">Array.&lt;String&gt;</span></span> | <span data-ttu-id="8bad4-p119">予定に必要なリソースを含む文字列の配列。配列は最大 100 エントリまでに制限されています。</span><span class="sxs-lookup"><span data-stu-id="8bad4-p119">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="8bad4-401">文字列</span><span class="sxs-lookup"><span data-stu-id="8bad4-401">String</span></span> | <span data-ttu-id="8bad4-p120">予定の件名を含む文字列。文字列は最大 255 文字までに制限されています。</span><span class="sxs-lookup"><span data-stu-id="8bad4-p120">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="8bad4-404">文字列</span><span class="sxs-lookup"><span data-stu-id="8bad4-404">String</span></span> | <span data-ttu-id="8bad4-p121">予定の本文。本文の内容は、最大サイズが 32 KB までに制限されています。</span><span class="sxs-lookup"><span data-stu-id="8bad4-p121">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="8bad4-407">要件</span><span class="sxs-lookup"><span data-stu-id="8bad4-407">Requirements</span></span>

|<span data-ttu-id="8bad4-408">要件</span><span class="sxs-lookup"><span data-stu-id="8bad4-408">Requirement</span></span>| <span data-ttu-id="8bad4-409">値</span><span class="sxs-lookup"><span data-stu-id="8bad4-409">Value</span></span>|
|---|---|
|[<span data-ttu-id="8bad4-410">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="8bad4-410">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8bad4-411">1.0</span><span class="sxs-lookup"><span data-stu-id="8bad4-411">1.0</span></span>|
|[<span data-ttu-id="8bad4-412">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="8bad4-412">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8bad4-413">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8bad4-413">ReadItem</span></span>|
|[<span data-ttu-id="8bad4-414">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8bad4-414">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8bad4-415">読み取り</span><span class="sxs-lookup"><span data-stu-id="8bad4-415">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8bad4-416">例</span><span class="sxs-lookup"><span data-stu-id="8bad4-416">Example</span></span>

```
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

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="8bad4-417">displayNewMessageForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="8bad4-417">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="8bad4-418">新しいメッセージを作成するためのフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="8bad4-418">Displays a form for creating a new message.</span></span>

<span data-ttu-id="8bad4-p122">`displayNewMessageForm` メソッドを使用すると、ユーザーが新しいメッセージを作成できるフォームが開きます。パラメータを指定すると、メッセージ フォーム フィールドにパラメータの内容が自動的に入力されます。</span><span class="sxs-lookup"><span data-stu-id="8bad4-p122">The `displayNewMessageForm` method opens a form that enables the user to create a new message. If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="8bad4-421">パラメータのいずれかが指定されたサイズ制限を超えた場合、または不明なパラメータ名が指定された場合には、例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="8bad4-421">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8bad4-422">パラメータ:</span><span class="sxs-lookup"><span data-stu-id="8bad4-422">Parameters:</span></span>

> [!NOTE]
> <span data-ttu-id="8bad4-423">すべてのパラメータは省略可能です。</span><span class="sxs-lookup"><span data-stu-id="8bad4-423">Note: All parameters are optional.</span></span>

|<span data-ttu-id="8bad4-424">名前</span><span class="sxs-lookup"><span data-stu-id="8bad4-424">Name</span></span>| <span data-ttu-id="8bad4-425">種類</span><span class="sxs-lookup"><span data-stu-id="8bad4-425">Type</span></span>| <span data-ttu-id="8bad4-426">説明</span><span class="sxs-lookup"><span data-stu-id="8bad4-426">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="8bad4-427">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="8bad4-427">Object</span></span> | <span data-ttu-id="8bad4-428">新しいメッセージを記述するパラメータの辞書。</span><span class="sxs-lookup"><span data-stu-id="8bad4-428">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="8bad4-429">配列。&lt;文字列&gt; | 配列。&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="8bad4-429">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="8bad4-p123">宛先行の各受信者の電子メール アドレスを含む文字列の配列、または `EmailAddressDetails` オブジェクトを含む配列。この配列は、最大 100 エントリまでに制限されています。</span><span class="sxs-lookup"><span data-stu-id="8bad4-p123">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="8bad4-432">配列。&lt;文字列&gt; | 配列。&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="8bad4-432">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="8bad4-p124">Cc 行の各受信者の電子メール アドレスを含む文字列の配列、または `EmailAddressDetails` オブジェクトを含む配列。この配列は、最大 100 エントリまでに制限されています。</span><span class="sxs-lookup"><span data-stu-id="8bad4-p124">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="8bad4-435">配列。&lt;文字列&gt; | 配列。&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="8bad4-435">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="8bad4-p125">Bcc 列の各受信者の電子メール アドレスを含む文字列の配列、または `EmailAddressDetails` オブジェクトを含む配列。この配列は、最大 100 エントリまでに制限されています。</span><span class="sxs-lookup"><span data-stu-id="8bad4-p125">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="8bad4-438">文字列</span><span class="sxs-lookup"><span data-stu-id="8bad4-438">String</span></span> | <span data-ttu-id="8bad4-p126">メッセージの件名を含む文字列。文字列は最大 255 文字までに制限されています。</span><span class="sxs-lookup"><span data-stu-id="8bad4-p126">A string containing the subject of the message. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="8bad4-441">文字列</span><span class="sxs-lookup"><span data-stu-id="8bad4-441">String</span></span> | <span data-ttu-id="8bad4-p127">メッセージの HTML 本文。本文の内容は、最大サイズが 32 KB までに制限されています。</span><span class="sxs-lookup"><span data-stu-id="8bad4-p127">The HTML body of the message. The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="8bad4-444">配列。&lt; オブジェクト&gt;</span><span class="sxs-lookup"><span data-stu-id="8bad4-444">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="8bad4-445">添付ファイルまたは添付アイテムである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="8bad4-445">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="8bad4-446">文字列</span><span class="sxs-lookup"><span data-stu-id="8bad4-446">String</span></span> | <span data-ttu-id="8bad4-p128">添付ファイルの種類を示します。添付ファイルの場合は`file`、添付項目の場合は`item`でなければなりません。</span><span class="sxs-lookup"><span data-stu-id="8bad4-p128">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="8bad4-449">文字列</span><span class="sxs-lookup"><span data-stu-id="8bad4-449">String</span></span> | <span data-ttu-id="8bad4-450">添付ファイル名を含む文字列で、255 文字以内で入力が可能です。</span><span class="sxs-lookup"><span data-stu-id="8bad4-450">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="8bad4-451">文字列</span><span class="sxs-lookup"><span data-stu-id="8bad4-451">String</span></span> | <span data-ttu-id="8bad4-p129">`type`が`file`に設定されている場合にのみ使用されます。ファイルの場所の URIです。</span><span class="sxs-lookup"><span data-stu-id="8bad4-p129">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="8bad4-454">ブール値</span><span class="sxs-lookup"><span data-stu-id="8bad4-454">Boolean</span></span> | <span data-ttu-id="8bad4-p130">`type`が`file`に設定されている場合にのみ使用されます。`true`の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="8bad4-p130">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="8bad4-457">文字列</span><span class="sxs-lookup"><span data-stu-id="8bad4-457">String</span></span> | <span data-ttu-id="8bad4-p131">`type` が `item` に設定されている場合にのみ使用されます。既存の電子メールの EWS アイテム ID です。最大 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="8bad4-p131">Only used if `type` is set to `item`. The EWS item id of the existing e-mail you want to attach to the new message. This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="8bad4-461">要件</span><span class="sxs-lookup"><span data-stu-id="8bad4-461">Requirements</span></span>

|<span data-ttu-id="8bad4-462">要件</span><span class="sxs-lookup"><span data-stu-id="8bad4-462">Requirement</span></span>| <span data-ttu-id="8bad4-463">値</span><span class="sxs-lookup"><span data-stu-id="8bad4-463">Value</span></span>|
|---|---|
|[<span data-ttu-id="8bad4-464">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="8bad4-464">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8bad4-465">1.6</span><span class="sxs-lookup"><span data-stu-id="8bad4-465">-16</span></span> |
|[<span data-ttu-id="8bad4-466">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="8bad4-466">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8bad4-467">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8bad4-467">ReadItem</span></span>|
|[<span data-ttu-id="8bad4-468">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8bad4-468">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8bad4-469">読み取り</span><span class="sxs-lookup"><span data-stu-id="8bad4-469">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8bad4-470">例</span><span class="sxs-lookup"><span data-stu-id="8bad4-470">Example</span></span>

```
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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="8bad4-471">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="8bad4-471">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="8bad4-472">REST API または Exchange Web サービスを呼び出すために使用するトークンを含む文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="8bad4-472">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="8bad4-p132">`getCallbackTokenAsync` メソッドは、非同期の呼び出しを行なって、ユーザーのメールボックスをホストする Exchange Server から opaque トークンを取得します。コールバック トークンの有効期間は 5 分です。</span><span class="sxs-lookup"><span data-stu-id="8bad4-p132">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="8bad4-475">可能な場合は常に、アドインでは Exchange Web サービスの代わりに REST API を使用することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="8bad4-475">Note: It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span> 

<span data-ttu-id="8bad4-476">**REST トークン**</span><span class="sxs-lookup"><span data-stu-id="8bad4-476">**REST Tokens**</span></span>

<span data-ttu-id="8bad4-p133">REST トークンが要求された場合 (`options.isRest = true`) には、作成されたトークンは Exchange Web サービスの呼び出しを認証するためには機能しません。このトークンは、アドインがマニフェストで [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) アクセス許可を指定しない限り、現在の項目およびその添付ファイルへの読み取り専用の範囲に制限されます。`ReadWriteMailbox` アクセス許可が指定された場合には、作成されるトークンは、メールを送信する機能など、メール、予定表、連絡先への読み取り/書き込みアクセスを付与します。</span><span class="sxs-lookup"><span data-stu-id="8bad4-p133">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="8bad4-480">アドインでは、`restUrl`プロパティを使用して、REST API 呼び出しを行うときに使用する正しい URL を決定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="8bad4-480">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="8bad4-481">**EWS トークン**</span><span class="sxs-lookup"><span data-stu-id="8bad4-481">**EWS Tokens**</span></span>

<span data-ttu-id="8bad4-p134">EWS トークンが要求された場合(`options.isRest = false`) には、作成されるトークンは REST API の呼び出しを認証するためには機能しません。このトークンは、現在の項目にアクセスできる範囲に制限されます。</span><span class="sxs-lookup"><span data-stu-id="8bad4-p134">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="8bad4-484">アドインでは、`ewsUrl` プロパティを使用して、EWS 呼び出しを行うときに使用する正しい URL を決定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="8bad4-484">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8bad4-485">パラメータ:</span><span class="sxs-lookup"><span data-stu-id="8bad4-485">Parameters:</span></span>

|<span data-ttu-id="8bad4-486">名前</span><span class="sxs-lookup"><span data-stu-id="8bad4-486">Name</span></span>| <span data-ttu-id="8bad4-487">種類</span><span class="sxs-lookup"><span data-stu-id="8bad4-487">Type</span></span>| <span data-ttu-id="8bad4-488">属性</span><span class="sxs-lookup"><span data-stu-id="8bad4-488">Attributes</span></span>| <span data-ttu-id="8bad4-489">説明</span><span class="sxs-lookup"><span data-stu-id="8bad4-489">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="8bad4-490">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="8bad4-490">Object</span></span> | <span data-ttu-id="8bad4-491">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="8bad4-491">&lt;optional&gt;</span></span> | <span data-ttu-id="8bad4-492">以下のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="8bad4-492">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="8bad4-493">ブール値</span><span class="sxs-lookup"><span data-stu-id="8bad4-493">Boolean</span></span> |  <span data-ttu-id="8bad4-494">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="8bad4-494">&lt;optional&gt;</span></span> | <span data-ttu-id="8bad4-p135">提供されたトークンを Outlook REST API または Exchange Web サービスに使用するかどうかを決定します。既定値は、`false`です。</span><span class="sxs-lookup"><span data-stu-id="8bad4-p135">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="8bad4-497">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="8bad4-497">Object</span></span> |  <span data-ttu-id="8bad4-498">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="8bad4-498">&lt;optional&gt;</span></span> | <span data-ttu-id="8bad4-499">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="8bad4-499">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="8bad4-500">関数</span><span class="sxs-lookup"><span data-stu-id="8bad4-500">function</span></span>||<span data-ttu-id="8bad4-p136">メソッドが完了すると、`callback` パラメータに渡された関数が、単一パラメータ `asyncResult` で呼び出されます。このパラメータは、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトです。トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="8bad4-p136">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8bad4-503">要件</span><span class="sxs-lookup"><span data-stu-id="8bad4-503">Requirements</span></span>

|<span data-ttu-id="8bad4-504">要件</span><span class="sxs-lookup"><span data-stu-id="8bad4-504">Requirement</span></span>| <span data-ttu-id="8bad4-505">値</span><span class="sxs-lookup"><span data-stu-id="8bad4-505">Value</span></span>|
|---|---|
|[<span data-ttu-id="8bad4-506">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="8bad4-506">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8bad4-507">1.5</span><span class="sxs-lookup"><span data-stu-id="8bad4-507">1.5</span></span> |
|[<span data-ttu-id="8bad4-508">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="8bad4-508">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8bad4-509">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8bad4-509">ReadItem</span></span>|
|[<span data-ttu-id="8bad4-510">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8bad4-510">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8bad4-511">新規作成と閲覧</span><span class="sxs-lookup"><span data-stu-id="8bad4-511">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="8bad4-512">例</span><span class="sxs-lookup"><span data-stu-id="8bad4-512">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="8bad4-513">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="8bad4-513">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="8bad4-514">Exchange Server から添付ファイルやアイテムを取得するために使用するトークンを含む文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="8bad4-514">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="8bad4-p137">`getCallbackTokenAsync` メソッドは、非同期の呼び出しを行なって、ユーザーのメールボックスをホストする Exchange Server から opaque トークンを取得します。コールバック トークンの有効期間は 5 分です。</span><span class="sxs-lookup"><span data-stu-id="8bad4-p137">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="8bad4-p138">このトークンと、添付ファイル識別子または項目識別子は、サードパーティーのシステムに渡すことができます。サードパーティーのシステムでは、添付ファイルまたは項目を返すための Exchange Web サービス (EWS) の [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) または [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) 操作を呼び出すのに、トークンをベアラー承認トークンとして使用します。たとえば、リモート サービスを作成して[選択した項目から添付ファイルを取得](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item)することができます。</span><span class="sxs-lookup"><span data-stu-id="8bad4-p138">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="8bad4-520">アプリでは、閲覧モードで `getCallbackTokenAsync` メソッドを呼び出すために、 **ReadItem** アクセス許可をアプリのマニフェストで指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="8bad4-520">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="8bad4-p139">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出して、`getCallbackTokenAsync` メソッドに渡すための項目識別子を取得する必要があります。アプリには、`saveAsync` メソッドを呼び出すために **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="8bad4-p139">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8bad4-523">パラメータ:</span><span class="sxs-lookup"><span data-stu-id="8bad4-523">Parameters:</span></span>

|<span data-ttu-id="8bad4-524">名前</span><span class="sxs-lookup"><span data-stu-id="8bad4-524">Name</span></span>| <span data-ttu-id="8bad4-525">種類</span><span class="sxs-lookup"><span data-stu-id="8bad4-525">Type</span></span>| <span data-ttu-id="8bad4-526">属性</span><span class="sxs-lookup"><span data-stu-id="8bad4-526">Attributes</span></span>| <span data-ttu-id="8bad4-527">説明</span><span class="sxs-lookup"><span data-stu-id="8bad4-527">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="8bad4-528">関数</span><span class="sxs-lookup"><span data-stu-id="8bad4-528">function</span></span>||<span data-ttu-id="8bad4-p140">メソッドが完了すると、`callback` パラメータに渡された関数が、単一パラメータ `asyncResult` で呼び出されます。このパラメータは、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトです。トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="8bad4-p140">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="8bad4-531">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="8bad4-531">Object</span></span>| <span data-ttu-id="8bad4-532">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="8bad4-532">&lt;optional&gt;</span></span>|<span data-ttu-id="8bad4-533">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="8bad4-533">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8bad4-534">要件</span><span class="sxs-lookup"><span data-stu-id="8bad4-534">Requirements</span></span>

|<span data-ttu-id="8bad4-535">要件</span><span class="sxs-lookup"><span data-stu-id="8bad4-535">Requirement</span></span>| <span data-ttu-id="8bad4-536">値</span><span class="sxs-lookup"><span data-stu-id="8bad4-536">Value</span></span>|
|---|---|
|[<span data-ttu-id="8bad4-537">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="8bad4-537">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8bad4-538">1.3</span><span class="sxs-lookup"><span data-stu-id="8bad4-538">1.3</span></span>|
|[<span data-ttu-id="8bad4-539">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="8bad4-539">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8bad4-540">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8bad4-540">ReadItem</span></span>|
|[<span data-ttu-id="8bad4-541">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8bad4-541">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8bad4-542">新規作成と閲覧</span><span class="sxs-lookup"><span data-stu-id="8bad4-542">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="8bad4-543">例</span><span class="sxs-lookup"><span data-stu-id="8bad4-543">Example</span></span>

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="8bad4-544">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="8bad4-544">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="8bad4-545">ユーザーと Office アドインを識別するトークンを取得します。</span><span class="sxs-lookup"><span data-stu-id="8bad4-545">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="8bad4-546">`getUserIdentityTokenAsync` メソッドは、[アドインとユーザーをサードパーティのシステムで識別して認証](https://docs.microsoft.com/outlook/add-ins/authentication)するのに使用できるトークンを返します。</span><span class="sxs-lookup"><span data-stu-id="8bad4-546">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](https://docs.microsoft.com/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="8bad4-547">パラメータ:</span><span class="sxs-lookup"><span data-stu-id="8bad4-547">Parameters:</span></span>

|<span data-ttu-id="8bad4-548">名前</span><span class="sxs-lookup"><span data-stu-id="8bad4-548">Name</span></span>| <span data-ttu-id="8bad4-549">種類</span><span class="sxs-lookup"><span data-stu-id="8bad4-549">Type</span></span>| <span data-ttu-id="8bad4-550">属性</span><span class="sxs-lookup"><span data-stu-id="8bad4-550">Attributes</span></span>| <span data-ttu-id="8bad4-551">説明</span><span class="sxs-lookup"><span data-stu-id="8bad4-551">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="8bad4-552">関数</span><span class="sxs-lookup"><span data-stu-id="8bad4-552">function</span></span>||<span data-ttu-id="8bad4-553">メソッドが完了すると、`callback` パラメータで渡された関数が、単一パラメータ `asyncResult` で呼び出されます。このパラメータは、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="8bad4-553">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="8bad4-554">トークンは、`asyncResult.value`プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="8bad4-554">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="8bad4-555">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="8bad4-555">Object</span></span>| <span data-ttu-id="8bad4-556">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="8bad4-556">&lt;optional&gt;</span></span>|<span data-ttu-id="8bad4-557">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="8bad4-557">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8bad4-558">要件</span><span class="sxs-lookup"><span data-stu-id="8bad4-558">Requirements</span></span>

|<span data-ttu-id="8bad4-559">要件</span><span class="sxs-lookup"><span data-stu-id="8bad4-559">Requirement</span></span>| <span data-ttu-id="8bad4-560">値</span><span class="sxs-lookup"><span data-stu-id="8bad4-560">Value</span></span>|
|---|---|
|[<span data-ttu-id="8bad4-561">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="8bad4-561">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8bad4-562">1.0</span><span class="sxs-lookup"><span data-stu-id="8bad4-562">1.0</span></span>|
|[<span data-ttu-id="8bad4-563">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="8bad4-563">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8bad4-564">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8bad4-564">ReadItem</span></span>|
|[<span data-ttu-id="8bad4-565">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8bad4-565">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8bad4-566">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="8bad4-566">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8bad4-567">例</span><span class="sxs-lookup"><span data-stu-id="8bad4-567">Example</span></span>

```js
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="8bad4-568">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="8bad4-568">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="8bad4-569">ユーザーのメールボックスをホストしている Exchange Server上の Exchange Web サービス (EWS) のサービスに対して非同期の要求を行います。</span><span class="sxs-lookup"><span data-stu-id="8bad4-569">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="8bad4-570">このメソッドは、次のシナリオではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="8bad4-570">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="8bad4-571">Outlook for iOS または Outlook for Android で</span><span class="sxs-lookup"><span data-stu-id="8bad4-571">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="8bad4-572">アドインが Gmail のメールボックスにロードされる場合</span><span class="sxs-lookup"><span data-stu-id="8bad4-572">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="8bad4-573">これらの場合、アドインではユーザーのメールボックスにアクセスするために、代わりに [ REST API を使用する](https://docs.microsoft.com/outlook/add-ins/use-rest-api)必要があります。</span><span class="sxs-lookup"><span data-stu-id="8bad4-573">In these cases, add-ins should [use REST APIs](https://docs.microsoft.com/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="8bad4-p141">`makeEwsRequestAsync` メソッドは、アドインの代わりに Exchange に EWS 要求を送信します。サポートされている EWS 操作の一覧については、「[ Outlook アドインから Web サービスを呼び出す](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="8bad4-p141">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange. See [Call web services from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="8bad4-576">`makeEwsRequestAsync` メソッドで、フォルダー関連アイテムを要求することはできません。</span><span class="sxs-lookup"><span data-stu-id="8bad4-576">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="8bad4-577">XML 要求では、UTF-8 エンコードを指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="8bad4-577">The XML request must specify UTF-8 encoding.</span></span>

```
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="8bad4-p142">アドインには、`makeEwsRequestAsync` メソッドを使用するために **ReadWriteMailbox** アクセス許可が必要です。**ReadWriteMailbox** アクセス許可と、`makeEwsRequestAsync` メソッドで呼び出すことのできる EWS 操作の使用の詳細については、「[ユーザーのメールボックスへのメール アドイン アクセスのアクセス許可を指定する](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="8bad4-p142">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="8bad4-580">サーバー管理者は、クライアント アクセス サーバー の EWS ディレクトリ上で `OAuthAuthentication` を true に設定して、`makeEwsRequestAsync` メソッドで EWS 要求を行えるようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="8bad4-580">NOTE: The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="8bad4-581">バージョンの相違点</span><span class="sxs-lookup"><span data-stu-id="8bad4-581">Version differences</span></span>

<span data-ttu-id="8bad4-582">バージョン 15.0.4535.1004 より前のバージョンの Outlook で実行しているメール アプリで `makeEwsRequestAsync` メソッドを使用する場合には、エンコード値を `ISO-8859-1` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="8bad4-582">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="8bad4-p143">メール アプリが Outlook on the web で実行されている場合には、エンコード値を設定する必要はありません。メールボックスを使用してメール アプリが Outlook で実行されているのか、Outlook on the web で実行されているのかを判断する必要があります。mailbox.diagnostics.hostVersion プロパティを使用すれば、どのバージョンの Outlook が実行されているのかがわかります。</span><span class="sxs-lookup"><span data-stu-id="8bad4-p143">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8bad4-586">パラメータ:</span><span class="sxs-lookup"><span data-stu-id="8bad4-586">Parameters:</span></span>

|<span data-ttu-id="8bad4-587">名前</span><span class="sxs-lookup"><span data-stu-id="8bad4-587">Name</span></span>| <span data-ttu-id="8bad4-588">種類</span><span class="sxs-lookup"><span data-stu-id="8bad4-588">Type</span></span>| <span data-ttu-id="8bad4-589">属性</span><span class="sxs-lookup"><span data-stu-id="8bad4-589">Attributes</span></span>| <span data-ttu-id="8bad4-590">説明</span><span class="sxs-lookup"><span data-stu-id="8bad4-590">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="8bad4-591">文字列</span><span class="sxs-lookup"><span data-stu-id="8bad4-591">String</span></span>||<span data-ttu-id="8bad4-592">EWS 要求。</span><span class="sxs-lookup"><span data-stu-id="8bad4-592">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="8bad4-593">関数</span><span class="sxs-lookup"><span data-stu-id="8bad4-593">function</span></span>||<span data-ttu-id="8bad4-594">メソッドが完了すると、`callback` パラメータで渡された関数が単一パラメータ `asyncResult` で呼び出されます。このパラメータは、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="8bad4-594">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="8bad4-p144">EWS 呼び出しの XML 結果は、`asyncResult.value` プロパティで文字列として提供されます。結果のサイズが 1 MB を超えている場合、エラー メッセージが返されます。</span><span class="sxs-lookup"><span data-stu-id="8bad4-p144">The XML result of the EWS call is provided as a string in the `asyncResult.value` property. If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="8bad4-597">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="8bad4-597">Object</span></span>| <span data-ttu-id="8bad4-598">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="8bad4-598">&lt;optional&gt;</span></span>|<span data-ttu-id="8bad4-599">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="8bad4-599">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8bad4-600">要件</span><span class="sxs-lookup"><span data-stu-id="8bad4-600">Requirements</span></span>

|<span data-ttu-id="8bad4-601">要件</span><span class="sxs-lookup"><span data-stu-id="8bad4-601">Requirement</span></span>| <span data-ttu-id="8bad4-602">値</span><span class="sxs-lookup"><span data-stu-id="8bad4-602">Value</span></span>|
|---|---|
|[<span data-ttu-id="8bad4-603">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="8bad4-603">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8bad4-604">1.0</span><span class="sxs-lookup"><span data-stu-id="8bad4-604">1.0</span></span>|
|[<span data-ttu-id="8bad4-605">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="8bad4-605">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8bad4-606">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="8bad4-606">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="8bad4-607">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8bad4-607">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8bad4-608">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="8bad4-608">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8bad4-609">例</span><span class="sxs-lookup"><span data-stu-id="8bad4-609">Example</span></span>

<span data-ttu-id="8bad4-610">次の例は、`makeEwsRequestAsync` を呼び出し、`GetItem` 操作を使って項目の件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="8bad4-610">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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