
# <a name="mailbox"></a><span data-ttu-id="0c407-101">mailbox</span><span class="sxs-lookup"><span data-stu-id="0c407-101">mailbox</span></span>

### <span data-ttu-id="0c407-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span><span class="sxs-lookup"><span data-stu-id="0c407-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span></span>

<span data-ttu-id="0c407-104">Microsoft Outlook と Microsoft Outlook on the web の Outlook アドイン オブジェクト モデルへのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="0c407-104">Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="0c407-105">要件</span><span class="sxs-lookup"><span data-stu-id="0c407-105">Requirements</span></span>

|<span data-ttu-id="0c407-106">要件</span><span class="sxs-lookup"><span data-stu-id="0c407-106">Requirement</span></span>| <span data-ttu-id="0c407-107">値</span><span class="sxs-lookup"><span data-stu-id="0c407-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="0c407-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0c407-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0c407-109">1.0</span><span class="sxs-lookup"><span data-stu-id="0c407-109">1.0</span></span>|
|[<span data-ttu-id="0c407-110">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0c407-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0c407-111">制限あり</span><span class="sxs-lookup"><span data-stu-id="0c407-111">Restricted</span></span>|
|[<span data-ttu-id="0c407-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0c407-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0c407-113">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="0c407-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="0c407-114">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="0c407-114">Members and methods</span></span>

| <span data-ttu-id="0c407-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="0c407-115">Member</span></span> | <span data-ttu-id="0c407-116">種類</span><span class="sxs-lookup"><span data-stu-id="0c407-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="0c407-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="0c407-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="0c407-118">メンバー</span><span class="sxs-lookup"><span data-stu-id="0c407-118">Member</span></span> |
| [<span data-ttu-id="0c407-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="0c407-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="0c407-120">メンバー</span><span class="sxs-lookup"><span data-stu-id="0c407-120">Member</span></span> |
| [<span data-ttu-id="0c407-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="0c407-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="0c407-122">メソッド</span><span class="sxs-lookup"><span data-stu-id="0c407-122">Method</span></span> |
| [<span data-ttu-id="0c407-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="0c407-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="0c407-124">メソッド</span><span class="sxs-lookup"><span data-stu-id="0c407-124">Method</span></span> |
| [<span data-ttu-id="0c407-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="0c407-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) | <span data-ttu-id="0c407-126">メソッド</span><span class="sxs-lookup"><span data-stu-id="0c407-126">Method</span></span> |
| [<span data-ttu-id="0c407-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="0c407-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="0c407-128">メソッド</span><span class="sxs-lookup"><span data-stu-id="0c407-128">Method</span></span> |
| [<span data-ttu-id="0c407-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="0c407-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="0c407-130">メソッド</span><span class="sxs-lookup"><span data-stu-id="0c407-130">Method</span></span> |
| [<span data-ttu-id="0c407-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="0c407-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="0c407-132">メソッド</span><span class="sxs-lookup"><span data-stu-id="0c407-132">Method</span></span> |
| [<span data-ttu-id="0c407-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="0c407-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="0c407-134">メソッド</span><span class="sxs-lookup"><span data-stu-id="0c407-134">Method</span></span> |
| [<span data-ttu-id="0c407-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="0c407-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="0c407-136">メソッド</span><span class="sxs-lookup"><span data-stu-id="0c407-136">Method</span></span> |
| [<span data-ttu-id="0c407-137">displayNewMessageForm</span><span class="sxs-lookup"><span data-stu-id="0c407-137">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="0c407-138">メソッド</span><span class="sxs-lookup"><span data-stu-id="0c407-138">Method</span></span> |
| [<span data-ttu-id="0c407-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="0c407-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="0c407-140">メソッド</span><span class="sxs-lookup"><span data-stu-id="0c407-140">Method</span></span> |
| [<span data-ttu-id="0c407-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="0c407-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="0c407-142">メソッド</span><span class="sxs-lookup"><span data-stu-id="0c407-142">Method</span></span> |
| [<span data-ttu-id="0c407-143">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="0c407-143">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="0c407-144">メソッド</span><span class="sxs-lookup"><span data-stu-id="0c407-144">Method</span></span> |
| [<span data-ttu-id="0c407-145">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="0c407-145">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="0c407-146">メソッド</span><span class="sxs-lookup"><span data-stu-id="0c407-146">Method</span></span> |
| [<span data-ttu-id="0c407-147">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="0c407-147">removeHandlerAsync</span></span>](#removehandlerasynceventtype-handler-options-callback) | <span data-ttu-id="0c407-148">メソッド</span><span class="sxs-lookup"><span data-stu-id="0c407-148">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="0c407-149">名前空間</span><span class="sxs-lookup"><span data-stu-id="0c407-149">Namespaces</span></span>

<span data-ttu-id="0c407-150">[diagnostics](Office.context.mailbox.diagnostics.md):Outlook アドインに診断情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="0c407-150">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="0c407-151">[item](Office.context.mailbox.item.md):Outlook アドインのメッセージや予定にアクセスするためのメソッドとプロパティを提供します。</span><span class="sxs-lookup"><span data-stu-id="0c407-151">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="0c407-152">[userProfile](Office.context.mailbox.userProfile.md):Outlook アドインのユーザーに関する情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="0c407-152">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="0c407-153">メンバー</span><span class="sxs-lookup"><span data-stu-id="0c407-153">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="0c407-154">ewsUrl :String</span><span class="sxs-lookup"><span data-stu-id="0c407-154">ewsUrl :String</span></span>

<span data-ttu-id="0c407-p102">このメール アカウントの Exchange Web サービス (EWS) エンドポイントの URL を取得します。読み取りモードのみです。</span><span class="sxs-lookup"><span data-stu-id="0c407-p102">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="0c407-157">このメンバーは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="0c407-157">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="0c407-p103">
  `ewsUrl\` 値は、リモート サービスで、ユーザーのメールボックスに EWS 呼び出しを行うために使うことができます。たとえば、[選択したアイテムから添付ファイルを取得する](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="0c407-p103">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="0c407-160">アプリが閲覧モードで `ewsUrl` メンバーを呼び出すには、アプリのマニフェスト内に **ReadItem** アクセス許可が指定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="0c407-160">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="0c407-p104">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出してから、`ewsUrl` メンバーを使用する必要があります。アプリには、`saveAsync` メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="0c407-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="0c407-163">型:</span><span class="sxs-lookup"><span data-stu-id="0c407-163">Type:</span></span>

*   <span data-ttu-id="0c407-164">String</span><span class="sxs-lookup"><span data-stu-id="0c407-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0c407-165">要件</span><span class="sxs-lookup"><span data-stu-id="0c407-165">Requirements</span></span>

|<span data-ttu-id="0c407-166">要件</span><span class="sxs-lookup"><span data-stu-id="0c407-166">Requirement</span></span>| <span data-ttu-id="0c407-167">値</span><span class="sxs-lookup"><span data-stu-id="0c407-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="0c407-168">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0c407-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0c407-169">1.0</span><span class="sxs-lookup"><span data-stu-id="0c407-169">1.0</span></span>|
|[<span data-ttu-id="0c407-170">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0c407-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0c407-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0c407-171">ReadItem</span></span>|
|[<span data-ttu-id="0c407-172">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0c407-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0c407-173">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="0c407-173">Compose or read</span></span>|

#### <a name="resturl-string"></a><span data-ttu-id="0c407-174">restUrl :String</span><span class="sxs-lookup"><span data-stu-id="0c407-174">restUrl :String</span></span>

<span data-ttu-id="0c407-175">この電子メール アカウントの REST エンドポイントの URL を取得します。</span><span class="sxs-lookup"><span data-stu-id="0c407-175">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="0c407-176">`restUrl` 値は、ユーザーのメールボックスに [REST API](https://docs.microsoft.com/outlook/rest/) 呼び出しを行うために使用できます。</span><span class="sxs-lookup"><span data-stu-id="0c407-176">The `restUrl` value can be used to make [REST API](https://docs.microsoft.com/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="0c407-177">アプリが閲覧モードで `restUrl` メンバーを呼び出すには、アプリのマニフェスト内に **ReadItem** アクセス許可が指定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="0c407-177">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="0c407-p105">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出してから、`restUrl` メンバーを使用する必要があります。アプリには、`saveAsync` メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="0c407-p105">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="0c407-180">型:</span><span class="sxs-lookup"><span data-stu-id="0c407-180">Type:</span></span>

*   <span data-ttu-id="0c407-181">String</span><span class="sxs-lookup"><span data-stu-id="0c407-181">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0c407-182">要件</span><span class="sxs-lookup"><span data-stu-id="0c407-182">Requirements</span></span>

|<span data-ttu-id="0c407-183">要件</span><span class="sxs-lookup"><span data-stu-id="0c407-183">Requirement</span></span>| <span data-ttu-id="0c407-184">値</span><span class="sxs-lookup"><span data-stu-id="0c407-184">Value</span></span>|
|---|---|
|[<span data-ttu-id="0c407-185">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0c407-185">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0c407-186">1.5</span><span class="sxs-lookup"><span data-stu-id="0c407-186">1.5</span></span> |
|[<span data-ttu-id="0c407-187">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0c407-187">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0c407-188">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0c407-188">ReadItem</span></span>|
|[<span data-ttu-id="0c407-189">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0c407-189">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0c407-190">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="0c407-190">Compose or read</span></span>|

### <a name="methods"></a><span data-ttu-id="0c407-191">メソッド</span><span class="sxs-lookup"><span data-stu-id="0c407-191">Methods</span></span>

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="0c407-192">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="0c407-192">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="0c407-193">サポートされているイベントのイベント ハンドラーを追加します。</span><span class="sxs-lookup"><span data-stu-id="0c407-193">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="0c407-194">現在サポートされている唯一のイベントの種類は `Office.EventType.ItemChanged` です。これはユーザーが新しいアイテムを選択したときに呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="0c407-194">Currently the only supported event type is `Office.EventType.ItemChanged`, which is invoked when the user selects a new item.</span></span> <span data-ttu-id="0c407-195">このイベントは、ピン留め可能な作業ウィンドウを実装するアドインで使用され、現在選択されているアイテムに基づいて作業ウィンドウ UI をアドインで更新できるようにします。</span><span class="sxs-lookup"><span data-stu-id="0c407-195">Currently the only supported event type is , which is invoked when the user selects a new item. This event is used by add-ins that implement a pinnable taskpane, and allows the add-in to refresh the taskpane UI based on the currently selected item.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0c407-196">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="0c407-196">Parameters:</span></span>

| <span data-ttu-id="0c407-197">名前</span><span class="sxs-lookup"><span data-stu-id="0c407-197">Name</span></span> | <span data-ttu-id="0c407-198">型</span><span class="sxs-lookup"><span data-stu-id="0c407-198">Type</span></span> | <span data-ttu-id="0c407-199">属性</span><span class="sxs-lookup"><span data-stu-id="0c407-199">Attributes</span></span> | <span data-ttu-id="0c407-200">説明</span><span class="sxs-lookup"><span data-stu-id="0c407-200">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="0c407-201">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="0c407-201">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="0c407-202">ハンドラーを呼び出す必要のあるイベント。</span><span class="sxs-lookup"><span data-stu-id="0c407-202">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="0c407-203">Function</span><span class="sxs-lookup"><span data-stu-id="0c407-203">Function</span></span> || <span data-ttu-id="0c407-p107">イベントを処理する関数。関数は、オブジェクト リテラルである単一パラメーターを受け入れる必要があります。パラメーターの `type` プロパティは、`addHandlerAsync` に渡される `eventType` パラメーターと一致します。</span><span class="sxs-lookup"><span data-stu-id="0c407-p107">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="0c407-207">Object</span><span class="sxs-lookup"><span data-stu-id="0c407-207">Object</span></span> | <span data-ttu-id="0c407-208">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0c407-208">&lt;optional&gt;</span></span> | <span data-ttu-id="0c407-209">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="0c407-209">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="0c407-210">Object</span><span class="sxs-lookup"><span data-stu-id="0c407-210">Object</span></span> | <span data-ttu-id="0c407-211">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0c407-211">&lt;optional&gt;</span></span> | <span data-ttu-id="0c407-212">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="0c407-212">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="0c407-213">function</span><span class="sxs-lookup"><span data-stu-id="0c407-213">function</span></span>| <span data-ttu-id="0c407-214">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0c407-214">&lt;optional&gt;</span></span>|<span data-ttu-id="0c407-215">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="0c407-215">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0c407-216">要件</span><span class="sxs-lookup"><span data-stu-id="0c407-216">Requirements</span></span>

|<span data-ttu-id="0c407-217">要件</span><span class="sxs-lookup"><span data-stu-id="0c407-217">Requirement</span></span>| <span data-ttu-id="0c407-218">値</span><span class="sxs-lookup"><span data-stu-id="0c407-218">Value</span></span>|
|---|---|
|[<span data-ttu-id="0c407-219">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0c407-219">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0c407-220">1.5</span><span class="sxs-lookup"><span data-stu-id="0c407-220">1.5</span></span> |
|[<span data-ttu-id="0c407-221">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0c407-221">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0c407-222">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0c407-222">ReadItem</span></span> |
|[<span data-ttu-id="0c407-223">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0c407-223">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0c407-224">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="0c407-224">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="0c407-225">例</span><span class="sxs-lookup"><span data-stu-id="0c407-225">Example</span></span>

```js
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

####  <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="0c407-226">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="0c407-226">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="0c407-227">REST 形式のアイテム ID を EWS 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="0c407-227">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="0c407-228">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="0c407-228">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="0c407-p108">REST API ([Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) や [Microsoft Graph](https://graph.microsoft.io/) など) で取得されたアイテム ID は、Exchange Web サービス (EWS) に使用される形式とは異なる形式を使用します。`convertToEwsId` メソッドは、REST 形式の ID を EWS 用の適切な形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="0c407-p108">Item IDs retrieved via a REST API (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0c407-231">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="0c407-231">Parameters:</span></span>

|<span data-ttu-id="0c407-232">名前</span><span class="sxs-lookup"><span data-stu-id="0c407-232">Name</span></span>| <span data-ttu-id="0c407-233">型</span><span class="sxs-lookup"><span data-stu-id="0c407-233">Type</span></span>| <span data-ttu-id="0c407-234">説明</span><span class="sxs-lookup"><span data-stu-id="0c407-234">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="0c407-235">String</span><span class="sxs-lookup"><span data-stu-id="0c407-235">String</span></span>|<span data-ttu-id="0c407-236">Outlook REST API 形式のアイテム ID</span><span class="sxs-lookup"><span data-stu-id="0c407-236">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="0c407-237">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="0c407-237">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.restversion)|<span data-ttu-id="0c407-238">アイテム ID の取得に使用された Outlook REST API のバージョンを示す値。</span><span class="sxs-lookup"><span data-stu-id="0c407-238">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0c407-239">要件</span><span class="sxs-lookup"><span data-stu-id="0c407-239">Requirements</span></span>

|<span data-ttu-id="0c407-240">要件</span><span class="sxs-lookup"><span data-stu-id="0c407-240">Requirement</span></span>| <span data-ttu-id="0c407-241">値</span><span class="sxs-lookup"><span data-stu-id="0c407-241">Value</span></span>|
|---|---|
|[<span data-ttu-id="0c407-242">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0c407-242">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0c407-243">1.3</span><span class="sxs-lookup"><span data-stu-id="0c407-243">1.3</span></span>|
|[<span data-ttu-id="0c407-244">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0c407-244">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0c407-245">制限あり</span><span class="sxs-lookup"><span data-stu-id="0c407-245">Restricted</span></span>|
|[<span data-ttu-id="0c407-246">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0c407-246">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0c407-247">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="0c407-247">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0c407-248">戻り値:</span><span class="sxs-lookup"><span data-stu-id="0c407-248">Returns:</span></span>

<span data-ttu-id="0c407-249">型:String</span><span class="sxs-lookup"><span data-stu-id="0c407-249">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="0c407-250">例</span><span class="sxs-lookup"><span data-stu-id="0c407-250">Example</span></span>

```js
// Get an item's ID from a REST API
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the
// Outlook Mail API
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime"></a><span data-ttu-id="0c407-251">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="0c407-251">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)}</span></span>

<span data-ttu-id="0c407-252">クライアントのローカル時間で時間情報が含まれている辞書を取得します。</span><span class="sxs-lookup"><span data-stu-id="0c407-252">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="0c407-p109">Outlook 用メール アプリや Outlook Web App で使う日付と時刻では、異なるタイム ゾーンを使うことができます。Outlook では、クライアント コンピューターのタイム ゾーンを使います。Outlook Web App では、Exchange 管理センター (EAC) で設定されたタイム ゾーンを使います。ユーザー インターフェイスに表示される値が、常にユーザーが期待するタイム ゾーンと一致するように日付と時刻の値を処理する必要があります。</span><span class="sxs-lookup"><span data-stu-id="0c407-p109">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="0c407-p110">Outlook でメール アプリが実行されている場合、`convertToLocalClientTime` メソッドは、クライアント コンピューターのタイム ゾーンに設定された値のディクショナリ オブジェクトを返します。Outlook Web Apps でメール アプリが実行されている場合、`convertToLocalClientTime` メソッドは、EAC に指定されたタイム ゾーンに設定された値のディクショナリ オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="0c407-p110">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0c407-258">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="0c407-258">Parameters:</span></span>

|<span data-ttu-id="0c407-259">名前</span><span class="sxs-lookup"><span data-stu-id="0c407-259">Name</span></span>| <span data-ttu-id="0c407-260">型</span><span class="sxs-lookup"><span data-stu-id="0c407-260">Type</span></span>| <span data-ttu-id="0c407-261">説明</span><span class="sxs-lookup"><span data-stu-id="0c407-261">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="0c407-262">Date</span><span class="sxs-lookup"><span data-stu-id="0c407-262">Date</span></span>|<span data-ttu-id="0c407-263">日付オブジェクト</span><span class="sxs-lookup"><span data-stu-id="0c407-263">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0c407-264">要件</span><span class="sxs-lookup"><span data-stu-id="0c407-264">Requirements</span></span>

|<span data-ttu-id="0c407-265">要件</span><span class="sxs-lookup"><span data-stu-id="0c407-265">Requirement</span></span>| <span data-ttu-id="0c407-266">値</span><span class="sxs-lookup"><span data-stu-id="0c407-266">Value</span></span>|
|---|---|
|[<span data-ttu-id="0c407-267">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0c407-267">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0c407-268">1.0</span><span class="sxs-lookup"><span data-stu-id="0c407-268">1.0</span></span>|
|[<span data-ttu-id="0c407-269">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0c407-269">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0c407-270">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0c407-270">ReadItem</span></span>|
|[<span data-ttu-id="0c407-271">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0c407-271">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0c407-272">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="0c407-272">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0c407-273">戻り値:</span><span class="sxs-lookup"><span data-stu-id="0c407-273">Returns:</span></span>

<span data-ttu-id="0c407-274">型:[LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="0c407-274">Type: [LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)</span></span>

####  <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="0c407-275">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="0c407-275">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="0c407-276">EWS 形式のアイテム ID を REST 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="0c407-276">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="0c407-277">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="0c407-277">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="0c407-p111">EWS または `itemId` プロパティで取得されるアイテム ID は、REST API ([Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) や [Microsoft Graph](https://graph.microsoft.io/) など) に使用される形式とは異なる形式を使用します。`convertToRestId` メソッドは、EWS 形式の ID を REST 用の適切な形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="0c407-p111">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0c407-280">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="0c407-280">Parameters:</span></span>

|<span data-ttu-id="0c407-281">名前</span><span class="sxs-lookup"><span data-stu-id="0c407-281">Name</span></span>| <span data-ttu-id="0c407-282">型</span><span class="sxs-lookup"><span data-stu-id="0c407-282">Type</span></span>| <span data-ttu-id="0c407-283">説明</span><span class="sxs-lookup"><span data-stu-id="0c407-283">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="0c407-284">String</span><span class="sxs-lookup"><span data-stu-id="0c407-284">String</span></span>|<span data-ttu-id="0c407-285">Exchange Web サービス (EWS) 形式のアイテム ID</span><span class="sxs-lookup"><span data-stu-id="0c407-285">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="0c407-286">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="0c407-286">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.restversion)|<span data-ttu-id="0c407-287">変換後の ID を使用する Outlook REST API のバージョンを示す値。</span><span class="sxs-lookup"><span data-stu-id="0c407-287">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0c407-288">要件</span><span class="sxs-lookup"><span data-stu-id="0c407-288">Requirements</span></span>

|<span data-ttu-id="0c407-289">要件</span><span class="sxs-lookup"><span data-stu-id="0c407-289">Requirement</span></span>| <span data-ttu-id="0c407-290">値</span><span class="sxs-lookup"><span data-stu-id="0c407-290">Value</span></span>|
|---|---|
|[<span data-ttu-id="0c407-291">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0c407-291">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0c407-292">1.3</span><span class="sxs-lookup"><span data-stu-id="0c407-292">1.3</span></span>|
|[<span data-ttu-id="0c407-293">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0c407-293">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0c407-294">制限あり</span><span class="sxs-lookup"><span data-stu-id="0c407-294">Restricted</span></span>|
|[<span data-ttu-id="0c407-295">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0c407-295">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0c407-296">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="0c407-296">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0c407-297">戻り値:</span><span class="sxs-lookup"><span data-stu-id="0c407-297">Returns:</span></span>

<span data-ttu-id="0c407-298">型:String</span><span class="sxs-lookup"><span data-stu-id="0c407-298">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="0c407-299">例</span><span class="sxs-lookup"><span data-stu-id="0c407-299">Example</span></span>

```js
// Get the currently selected item's ID
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the
// Outlook Mail API
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="0c407-300">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="0c407-300">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="0c407-301">時間情報が含まれているディクショナリから日付オブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="0c407-301">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="0c407-302">`convertToUtcClientTime` メソッドは、ローカルの日付と時刻を含むディクショナリを、ローカルの日付と時刻の正しい値を持つ日付オブジェクトに変換します。</span><span class="sxs-lookup"><span data-stu-id="0c407-302">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0c407-303">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="0c407-303">Parameters:</span></span>

|<span data-ttu-id="0c407-304">名前</span><span class="sxs-lookup"><span data-stu-id="0c407-304">Name</span></span>| <span data-ttu-id="0c407-305">型</span><span class="sxs-lookup"><span data-stu-id="0c407-305">Type</span></span>| <span data-ttu-id="0c407-306">説明</span><span class="sxs-lookup"><span data-stu-id="0c407-306">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="0c407-307">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="0c407-307">LocalClientTime</span></span>](/javascript/api/outlook_1_6/office.LocalClientTime)|<span data-ttu-id="0c407-308">変換するローカル時刻の値。</span><span class="sxs-lookup"><span data-stu-id="0c407-308">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0c407-309">要件</span><span class="sxs-lookup"><span data-stu-id="0c407-309">Requirements</span></span>

|<span data-ttu-id="0c407-310">要件</span><span class="sxs-lookup"><span data-stu-id="0c407-310">Requirement</span></span>| <span data-ttu-id="0c407-311">値</span><span class="sxs-lookup"><span data-stu-id="0c407-311">Value</span></span>|
|---|---|
|[<span data-ttu-id="0c407-312">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0c407-312">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0c407-313">1.0</span><span class="sxs-lookup"><span data-stu-id="0c407-313">1.0</span></span>|
|[<span data-ttu-id="0c407-314">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0c407-314">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0c407-315">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0c407-315">ReadItem</span></span>|
|[<span data-ttu-id="0c407-316">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0c407-316">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0c407-317">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="0c407-317">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0c407-318">戻り値:</span><span class="sxs-lookup"><span data-stu-id="0c407-318">Returns:</span></span>

<span data-ttu-id="0c407-319">時間が UTC で表現された日付オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="0c407-319">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="0c407-320">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="0c407-320">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="0c407-321">Date</span><span class="sxs-lookup"><span data-stu-id="0c407-321">Date</span></span></dd>

</dl>

####  <a name="displayappointmentformitemid"></a><span data-ttu-id="0c407-322">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="0c407-322">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="0c407-323">既存の予定を表示します。</span><span class="sxs-lookup"><span data-stu-id="0c407-323">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="0c407-324">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="0c407-324">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="0c407-325">`displayAppointmentForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存の予定を開きます。</span><span class="sxs-lookup"><span data-stu-id="0c407-325">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="0c407-p112">Outlook for Mac では、この方法を使って、定期的な系列の一部ではない単一の予定や定期的な系列のマスター予定を表示できます。ただし、系列のインスタンスは表示できません。これは、Outlook for Mac においては定期的な系列のインスタンスのプロパティ (アイテム ID を含む) にアクセスできないためです。</span><span class="sxs-lookup"><span data-stu-id="0c407-p112">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="0c407-328">Outlook Web App では、このメソッドは指定されたフォームの本文が 32 KB 以下の文字数の場合にフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="0c407-328">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="0c407-329">指定のアイテム識別子が既存の予定を表していない場合は、クライアント コンピューターまたはデバイスで空のウィンドウが開き、エラー メッセージは返されません。</span><span class="sxs-lookup"><span data-stu-id="0c407-329">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0c407-330">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="0c407-330">Parameters:</span></span>

|<span data-ttu-id="0c407-331">名前</span><span class="sxs-lookup"><span data-stu-id="0c407-331">Name</span></span>| <span data-ttu-id="0c407-332">型</span><span class="sxs-lookup"><span data-stu-id="0c407-332">Type</span></span>| <span data-ttu-id="0c407-333">説明</span><span class="sxs-lookup"><span data-stu-id="0c407-333">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="0c407-334">String</span><span class="sxs-lookup"><span data-stu-id="0c407-334">String</span></span>|<span data-ttu-id="0c407-335">既存の予定の Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="0c407-335">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0c407-336">要件</span><span class="sxs-lookup"><span data-stu-id="0c407-336">Requirements</span></span>

|<span data-ttu-id="0c407-337">要件</span><span class="sxs-lookup"><span data-stu-id="0c407-337">Requirement</span></span>| <span data-ttu-id="0c407-338">値</span><span class="sxs-lookup"><span data-stu-id="0c407-338">Value</span></span>|
|---|---|
|[<span data-ttu-id="0c407-339">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0c407-339">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0c407-340">1.0</span><span class="sxs-lookup"><span data-stu-id="0c407-340">1.0</span></span>|
|[<span data-ttu-id="0c407-341">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0c407-341">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0c407-342">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0c407-342">ReadItem</span></span>|
|[<span data-ttu-id="0c407-343">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0c407-343">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0c407-344">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="0c407-344">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="0c407-345">例</span><span class="sxs-lookup"><span data-stu-id="0c407-345">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  <a name="displaymessageformitemid"></a><span data-ttu-id="0c407-346">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="0c407-346">displayMessageForm(itemId)</span></span>

<span data-ttu-id="0c407-347">既存のメッセージを表示します。</span><span class="sxs-lookup"><span data-stu-id="0c407-347">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="0c407-348">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="0c407-348">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="0c407-349">`displayMessageForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存のメッセージを開きます。</span><span class="sxs-lookup"><span data-stu-id="0c407-349">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="0c407-350">Outlook Web App では、このメソッドは指定されたフォームの本文が 32 KB 以下の文字数の場合にフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="0c407-350">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="0c407-351">指定のアイテム識別子が既存のメッセージを表していない場合は、クライアント コンピューターにはメッセージは表示されず、エラー メッセージも返されません。</span><span class="sxs-lookup"><span data-stu-id="0c407-351">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="0c407-p113">予定を表す `itemId` を含む `displayMessageForm` を使用しないでください。既存の予定を表示するには、`displayAppointmentForm` メソッドを使用します。新しい予定を作成するフォームを表示するには、`displayNewAppointmentForm` メソッドを使用します。</span><span class="sxs-lookup"><span data-stu-id="0c407-p113">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0c407-354">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="0c407-354">Parameters:</span></span>

|<span data-ttu-id="0c407-355">名前</span><span class="sxs-lookup"><span data-stu-id="0c407-355">Name</span></span>| <span data-ttu-id="0c407-356">型</span><span class="sxs-lookup"><span data-stu-id="0c407-356">Type</span></span>| <span data-ttu-id="0c407-357">説明</span><span class="sxs-lookup"><span data-stu-id="0c407-357">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="0c407-358">String</span><span class="sxs-lookup"><span data-stu-id="0c407-358">String</span></span>|<span data-ttu-id="0c407-359">既存のメッセージの Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="0c407-359">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0c407-360">要件</span><span class="sxs-lookup"><span data-stu-id="0c407-360">Requirements</span></span>

|<span data-ttu-id="0c407-361">要件</span><span class="sxs-lookup"><span data-stu-id="0c407-361">Requirement</span></span>| <span data-ttu-id="0c407-362">値</span><span class="sxs-lookup"><span data-stu-id="0c407-362">Value</span></span>|
|---|---|
|[<span data-ttu-id="0c407-363">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0c407-363">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0c407-364">1.0</span><span class="sxs-lookup"><span data-stu-id="0c407-364">1.0</span></span>|
|[<span data-ttu-id="0c407-365">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0c407-365">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0c407-366">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0c407-366">ReadItem</span></span>|
|[<span data-ttu-id="0c407-367">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0c407-367">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0c407-368">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="0c407-368">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="0c407-369">例</span><span class="sxs-lookup"><span data-stu-id="0c407-369">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="0c407-370">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="0c407-370">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="0c407-371">新しい予定を作成するためのフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="0c407-371">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="0c407-372">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="0c407-372">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="0c407-p114">`displayNewAppointmentForm` メソッドを使用すると、ユーザーが新しい予定または会議を作成できるフォームが開きます。パラメーターを指定すると、予定のフォーム フィールドにパラメーターの内容が自動的に設定されます。</span><span class="sxs-lookup"><span data-stu-id="0c407-p114">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="0c407-p115">このメソッドは、Outlook Web App と OWA for Devices において、出席者フィールドが含まれるフォームを必ず表示します。入力引数として出席者を指定しないと、このメソッドにより **[保存]** ボタンのあるフォームが表示されます。出席者を指定した場合には、フォームにその出席者と **[送信]** ボタンが表示されます。</span><span class="sxs-lookup"><span data-stu-id="0c407-p115">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="0c407-p116">Outlook リッチ クライアントと Outlook RT で、`requiredAttendees`、`optionalAttendees`、または `resources` パラメーターに出席者またはリソースを指定し、このメソッドを実行すると、**[送信]** ボタンがある会議フォームが表示されます。受信者を指定せずにこのメソッドを実行すると、**[保存して閉じる]** ボタンがある予定フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="0c407-p116">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="0c407-380">パラメーターのいずれかが指定のサイズ制限を超える場合、または不明なパラメーター名が指定されている場合は、例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="0c407-380">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0c407-381">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="0c407-381">Parameters:</span></span>

> [!NOTE]
> <span data-ttu-id="0c407-382">すべてのパラメーターは省略可能です。</span><span class="sxs-lookup"><span data-stu-id="0c407-382">All parameters are optional.</span></span>

|<span data-ttu-id="0c407-383">名前</span><span class="sxs-lookup"><span data-stu-id="0c407-383">Name</span></span>| <span data-ttu-id="0c407-384">型</span><span class="sxs-lookup"><span data-stu-id="0c407-384">Type</span></span>| <span data-ttu-id="0c407-385">説明</span><span class="sxs-lookup"><span data-stu-id="0c407-385">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="0c407-386">Object</span><span class="sxs-lookup"><span data-stu-id="0c407-386">Object</span></span> | <span data-ttu-id="0c407-387">新しい予定を記述するパラメーターのディクショナリ。</span><span class="sxs-lookup"><span data-stu-id="0c407-387">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="0c407-388">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="0c407-388">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="0c407-p117">予定に必要な各出席者について、メール アドレスを含む文字列の配列、または `EmailAddressDetails` オブジェクトを含む配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="0c407-p117">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="0c407-391">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="0c407-391">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="0c407-p118">予定の各任意出席者について、メール アドレスを含む文字列の配列、または `EmailAddressDetails` オブジェクトを含む配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="0c407-p118">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="0c407-394">日付</span><span class="sxs-lookup"><span data-stu-id="0c407-394">Date</span></span> | <span data-ttu-id="0c407-395">予定の開始日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="0c407-395">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="0c407-396">Date</span><span class="sxs-lookup"><span data-stu-id="0c407-396">Date</span></span> | <span data-ttu-id="0c407-397">予定の終了日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="0c407-397">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="0c407-398">String</span><span class="sxs-lookup"><span data-stu-id="0c407-398">String</span></span> | <span data-ttu-id="0c407-p119">予定の場所を含む文字列。文字列は最大 255 文字に制限されます。</span><span class="sxs-lookup"><span data-stu-id="0c407-p119">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="0c407-401">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="0c407-401">Array.&lt;String&gt;</span></span> | <span data-ttu-id="0c407-p120">予定に必要なリソースを含む文字列の配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="0c407-p120">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="0c407-404">String</span><span class="sxs-lookup"><span data-stu-id="0c407-404">String</span></span> | <span data-ttu-id="0c407-p121">予定の件名を含む文字列です。文字列は最大 255 文字に制限されます。</span><span class="sxs-lookup"><span data-stu-id="0c407-p121">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="0c407-407">String</span><span class="sxs-lookup"><span data-stu-id="0c407-407">String</span></span> | <span data-ttu-id="0c407-p122">予定の本文。本文の内容は、最大サイズが 32 KB に制限されます。</span><span class="sxs-lookup"><span data-stu-id="0c407-p122">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="0c407-410">要件</span><span class="sxs-lookup"><span data-stu-id="0c407-410">Requirements</span></span>

|<span data-ttu-id="0c407-411">要件</span><span class="sxs-lookup"><span data-stu-id="0c407-411">Requirement</span></span>| <span data-ttu-id="0c407-412">値</span><span class="sxs-lookup"><span data-stu-id="0c407-412">Value</span></span>|
|---|---|
|[<span data-ttu-id="0c407-413">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0c407-413">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0c407-414">1.0</span><span class="sxs-lookup"><span data-stu-id="0c407-414">1.0</span></span>|
|[<span data-ttu-id="0c407-415">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0c407-415">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0c407-416">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0c407-416">ReadItem</span></span>|
|[<span data-ttu-id="0c407-417">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0c407-417">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0c407-418">読み取り</span><span class="sxs-lookup"><span data-stu-id="0c407-418">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0c407-419">例</span><span class="sxs-lookup"><span data-stu-id="0c407-419">Example</span></span>

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

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="0c407-420">displayNewMessageForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="0c407-420">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="0c407-421">新しいメッセージを作成するためのフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="0c407-421">Displays a form for creating a new message.</span></span>

<span data-ttu-id="0c407-422">`displayNewMessageForm` メソッドは、ユーザーが新しいメッセージを作成できるようにするフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="0c407-422">The `displayNewMessageForm` method opens a form that enables the user to create a new message.</span></span> <span data-ttu-id="0c407-423">パラメーターを指定すると、メッセージ フォーム フィールドにはパラメーターのコンテンツが自動的に入力されます。</span><span class="sxs-lookup"><span data-stu-id="0c407-423">If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="0c407-424">パラメーターのいずれかが指定のサイズ制限を超える場合、または不明なパラメーター名が指定されている場合は、例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="0c407-424">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0c407-425">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="0c407-425">Parameters:</span></span>

> [!NOTE]
> <span data-ttu-id="0c407-426">すべてのパラメーターは省略可能です。</span><span class="sxs-lookup"><span data-stu-id="0c407-426">All parameters are optional.</span></span>

|<span data-ttu-id="0c407-427">名前</span><span class="sxs-lookup"><span data-stu-id="0c407-427">Name</span></span>| <span data-ttu-id="0c407-428">型</span><span class="sxs-lookup"><span data-stu-id="0c407-428">Type</span></span>| <span data-ttu-id="0c407-429">説明</span><span class="sxs-lookup"><span data-stu-id="0c407-429">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="0c407-430">Object</span><span class="sxs-lookup"><span data-stu-id="0c407-430">Object</span></span> | <span data-ttu-id="0c407-431">新しいメッセージを記述するパラメーターのディクショナリ。</span><span class="sxs-lookup"><span data-stu-id="0c407-431">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="0c407-432">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="0c407-432">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="0c407-433">メール アドレスを含む文字列の配列、または To 行の各受信者の `EmailAddressDetails` オブジェクトを含む配列。</span><span class="sxs-lookup"><span data-stu-id="0c407-433">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line.</span></span> <span data-ttu-id="0c407-434">配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="0c407-434">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="0c407-435">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="0c407-435">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="0c407-436">メール アドレスを含む文字列の配列、または Cc 行の各受信者の `EmailAddressDetails` オブジェクトを含む配列。</span><span class="sxs-lookup"><span data-stu-id="0c407-436">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line.</span></span> <span data-ttu-id="0c407-437">配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="0c407-437">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="0c407-438">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="0c407-438">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="0c407-439">メール アドレスを含む文字列の配列、または Bcc 行の各受信者の `EmailAddressDetails` オブジェクトを含む配列。</span><span class="sxs-lookup"><span data-stu-id="0c407-439">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line.</span></span> <span data-ttu-id="0c407-440">配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="0c407-440">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="0c407-441">String</span><span class="sxs-lookup"><span data-stu-id="0c407-441">String</span></span> | <span data-ttu-id="0c407-442">メッセージの件名を含む文字列。</span><span class="sxs-lookup"><span data-stu-id="0c407-442">A string containing the subject of the message.</span></span> <span data-ttu-id="0c407-443">文字列は最大 255 文字に制限されます。</span><span class="sxs-lookup"><span data-stu-id="0c407-443">The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="0c407-444">String</span><span class="sxs-lookup"><span data-stu-id="0c407-444">String</span></span> | <span data-ttu-id="0c407-445">メッセージの HTML 本文。</span><span class="sxs-lookup"><span data-stu-id="0c407-445">The HTML body of the message.</span></span> <span data-ttu-id="0c407-446">本文の内容は、最大サイズが 32 KB に制限されます。</span><span class="sxs-lookup"><span data-stu-id="0c407-446">The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="0c407-447">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="0c407-447">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="0c407-448">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="0c407-448">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="0c407-449">String</span><span class="sxs-lookup"><span data-stu-id="0c407-449">String</span></span> | <span data-ttu-id="0c407-p129">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="0c407-p129">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="0c407-452">String</span><span class="sxs-lookup"><span data-stu-id="0c407-452">String</span></span> | <span data-ttu-id="0c407-453">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="0c407-453">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="0c407-454">String</span><span class="sxs-lookup"><span data-stu-id="0c407-454">String</span></span> | <span data-ttu-id="0c407-p130">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="0c407-p130">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="0c407-457">ブール値</span><span class="sxs-lookup"><span data-stu-id="0c407-457">Boolean</span></span> | <span data-ttu-id="0c407-p131">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="0c407-p131">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="0c407-460">String</span><span class="sxs-lookup"><span data-stu-id="0c407-460">String</span></span> | <span data-ttu-id="0c407-461">`type` が `item` に設定されている場合にのみ使用されます。</span><span class="sxs-lookup"><span data-stu-id="0c407-461">Only used if `type` is set to `item`.</span></span> <span data-ttu-id="0c407-462">新しいメッセージに添付する必要がある既存の電子メールの EWS のアイテム ID です。</span><span class="sxs-lookup"><span data-stu-id="0c407-462">The EWS item id of the existing e-mail you want to attach to the new message.</span></span> <span data-ttu-id="0c407-463">最大 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="0c407-463">This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="0c407-464">要件</span><span class="sxs-lookup"><span data-stu-id="0c407-464">Requirements</span></span>

|<span data-ttu-id="0c407-465">要件</span><span class="sxs-lookup"><span data-stu-id="0c407-465">Requirement</span></span>| <span data-ttu-id="0c407-466">値</span><span class="sxs-lookup"><span data-stu-id="0c407-466">Value</span></span>|
|---|---|
|[<span data-ttu-id="0c407-467">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0c407-467">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0c407-468">1.6</span><span class="sxs-lookup"><span data-stu-id="0c407-468">1.6</span></span> |
|[<span data-ttu-id="0c407-469">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0c407-469">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0c407-470">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0c407-470">ReadItem</span></span>|
|[<span data-ttu-id="0c407-471">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0c407-471">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0c407-472">読み取り</span><span class="sxs-lookup"><span data-stu-id="0c407-472">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0c407-473">例</span><span class="sxs-lookup"><span data-stu-id="0c407-473">Example</span></span>

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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="0c407-474">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="0c407-474">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="0c407-475">REST API または Exchange Web サービスを呼び出すために使用するトークンを含む文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="0c407-475">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="0c407-p133">`getCallbackTokenAsync` メソッドは、ユーザーのメールボックスをホストする Exchange Server から不透明なトークンを取得する非同期の呼び出しを行います。コールバック トークンの有効期間は 5 分です。</span><span class="sxs-lookup"><span data-stu-id="0c407-p133">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="0c407-478">可能であれば、アドインでは Exchange Web サービスの代わりに REST API を使用することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="0c407-478">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span> 

<span data-ttu-id="0c407-479">**REST トークン**</span><span class="sxs-lookup"><span data-stu-id="0c407-479">**REST Tokens**</span></span>

<span data-ttu-id="0c407-p134">REST トークンが要求された場合 (`options.isRest = true`)、結果トークンは Exchange Web サービスの呼び出しを認証するためには機能しません。アドインがマニフェストで [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) アクセス許可を指定していない限り、トークンの範囲は現在のアイテムとその添付ファイルへの読み取り専用アクセスに制限されます。`ReadWriteMailbox` アクセス許可が指定されている場合は、結果トークンは、メールを送信する機能など、メール、カレンダー、連絡先への読み取り/書き込みアクセスを付与します。</span><span class="sxs-lookup"><span data-stu-id="0c407-p134">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="0c407-483">アドインでは、`restUrl` プロパティを使用して、REST API 呼び出しを行うときに使用する正しい URL を決定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="0c407-483">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="0c407-484">**EWS トークン**</span><span class="sxs-lookup"><span data-stu-id="0c407-484">**EWS Tokens**</span></span>

<span data-ttu-id="0c407-p135">EWS トークンが要求された場合 (`options.isRest = false`)、結果トークンは REST API 呼び出しを認証するためには機能しません。トークンの範囲は、現在のアイテムへのアクセスに制限されます。</span><span class="sxs-lookup"><span data-stu-id="0c407-p135">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="0c407-487">アドインでは、`ewsUrl` プロパティを使用して、EWS 呼び出しを行うときに使用する正しい URL を決定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="0c407-487">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0c407-488">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="0c407-488">Parameters:</span></span>

|<span data-ttu-id="0c407-489">名前</span><span class="sxs-lookup"><span data-stu-id="0c407-489">Name</span></span>| <span data-ttu-id="0c407-490">型</span><span class="sxs-lookup"><span data-stu-id="0c407-490">Type</span></span>| <span data-ttu-id="0c407-491">属性</span><span class="sxs-lookup"><span data-stu-id="0c407-491">Attributes</span></span>| <span data-ttu-id="0c407-492">説明</span><span class="sxs-lookup"><span data-stu-id="0c407-492">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="0c407-493">Object</span><span class="sxs-lookup"><span data-stu-id="0c407-493">Object</span></span> | <span data-ttu-id="0c407-494">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0c407-494">&lt;optional&gt;</span></span> | <span data-ttu-id="0c407-495">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="0c407-495">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="0c407-496">Boolean</span><span class="sxs-lookup"><span data-stu-id="0c407-496">Boolean</span></span> |  <span data-ttu-id="0c407-497">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="0c407-497">&lt;optional&gt;</span></span> | <span data-ttu-id="0c407-p136">提供されたトークンを Outlook REST API または Exchange Web サービスに使用するかどうかを決定します。既定値は、`false` です。</span><span class="sxs-lookup"><span data-stu-id="0c407-p136">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="0c407-500">Object</span><span class="sxs-lookup"><span data-stu-id="0c407-500">Object</span></span> |  <span data-ttu-id="0c407-501">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="0c407-501">&lt;optional&gt;</span></span> | <span data-ttu-id="0c407-502">非同期メソッドに渡される状態データ。</span><span class="sxs-lookup"><span data-stu-id="0c407-502">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="0c407-503">function</span><span class="sxs-lookup"><span data-stu-id="0c407-503">function</span></span>||<span data-ttu-id="0c407-p137">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="0c407-p137">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0c407-506">要件</span><span class="sxs-lookup"><span data-stu-id="0c407-506">Requirements</span></span>

|<span data-ttu-id="0c407-507">要件</span><span class="sxs-lookup"><span data-stu-id="0c407-507">Requirement</span></span>| <span data-ttu-id="0c407-508">値</span><span class="sxs-lookup"><span data-stu-id="0c407-508">Value</span></span>|
|---|---|
|[<span data-ttu-id="0c407-509">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0c407-509">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0c407-510">1.5</span><span class="sxs-lookup"><span data-stu-id="0c407-510">1.5</span></span> |
|[<span data-ttu-id="0c407-511">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0c407-511">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0c407-512">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0c407-512">ReadItem</span></span>|
|[<span data-ttu-id="0c407-513">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0c407-513">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0c407-514">新規作成と閲覧</span><span class="sxs-lookup"><span data-stu-id="0c407-514">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="0c407-515">例</span><span class="sxs-lookup"><span data-stu-id="0c407-515">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="0c407-516">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="0c407-516">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="0c407-517">Exchange Server から添付ファイルやアイテムを取得するために使うトークンを含む文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="0c407-517">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="0c407-p138">`getCallbackTokenAsync` メソッドは、ユーザーのメールボックスをホストする Exchange Server から不透明なトークンを取得する非同期の呼び出しを行います。コールバック トークンの有効期間は 5 分です。</span><span class="sxs-lookup"><span data-stu-id="0c407-p138">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="0c407-p139">トークンと予定の識別子またはアイテムの識別子をサードパーティ システムに渡すことができます。サードパーティ システムは、トークンをベアラー承認トークンとして使用し、Exchange Web サービス (EWS) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) または [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) 操作を呼び出して、添付ファイルまたはアイテムを返します。たとえば、[選択したアイテムから添付ファイルを取得する](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="0c407-p139">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="0c407-523">アプリが閲覧モードで `getCallbackTokenAsync` メソッドを呼び出すには、アプリのマニフェスト内に **ReadItem** アクセス許可が指定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="0c407-523">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="0c407-p140">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出してアイテムの識別子を `getCallbackTokenAsync` メソッドに渡す必要があります。アプリには、`saveAsync` メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="0c407-p140">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0c407-526">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="0c407-526">Parameters:</span></span>

|<span data-ttu-id="0c407-527">名前</span><span class="sxs-lookup"><span data-stu-id="0c407-527">Name</span></span>| <span data-ttu-id="0c407-528">型</span><span class="sxs-lookup"><span data-stu-id="0c407-528">Type</span></span>| <span data-ttu-id="0c407-529">属性</span><span class="sxs-lookup"><span data-stu-id="0c407-529">Attributes</span></span>| <span data-ttu-id="0c407-530">説明</span><span class="sxs-lookup"><span data-stu-id="0c407-530">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="0c407-531">function</span><span class="sxs-lookup"><span data-stu-id="0c407-531">function</span></span>||<span data-ttu-id="0c407-p141">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="0c407-p141">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="0c407-534">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="0c407-534">Object</span></span>| <span data-ttu-id="0c407-535">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="0c407-535">&lt;optional&gt;</span></span>|<span data-ttu-id="0c407-536">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="0c407-536">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0c407-537">要件</span><span class="sxs-lookup"><span data-stu-id="0c407-537">Requirements</span></span>

|<span data-ttu-id="0c407-538">要件</span><span class="sxs-lookup"><span data-stu-id="0c407-538">Requirement</span></span>| <span data-ttu-id="0c407-539">値</span><span class="sxs-lookup"><span data-stu-id="0c407-539">Value</span></span>|
|---|---|
|[<span data-ttu-id="0c407-540">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0c407-540">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0c407-541">1.3</span><span class="sxs-lookup"><span data-stu-id="0c407-541">1.3</span></span>|
|[<span data-ttu-id="0c407-542">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0c407-542">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0c407-543">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0c407-543">ReadItem</span></span>|
|[<span data-ttu-id="0c407-544">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0c407-544">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0c407-545">新規作成と閲覧</span><span class="sxs-lookup"><span data-stu-id="0c407-545">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="0c407-546">例</span><span class="sxs-lookup"><span data-stu-id="0c407-546">Example</span></span>

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="0c407-547">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="0c407-547">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="0c407-548">ユーザーと Office アドインを識別するトークンを取得します。</span><span class="sxs-lookup"><span data-stu-id="0c407-548">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="0c407-549">`getUserIdentityTokenAsync` メソッドは、[アドインとユーザーをサード パーティのシステムで識別して認証](https://docs.microsoft.com/outlook/add-ins/authentication)することのできるトークンを返します。</span><span class="sxs-lookup"><span data-stu-id="0c407-549">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](https://docs.microsoft.com/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="0c407-550">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="0c407-550">Parameters:</span></span>

|<span data-ttu-id="0c407-551">名前</span><span class="sxs-lookup"><span data-stu-id="0c407-551">Name</span></span>| <span data-ttu-id="0c407-552">型</span><span class="sxs-lookup"><span data-stu-id="0c407-552">Type</span></span>| <span data-ttu-id="0c407-553">属性</span><span class="sxs-lookup"><span data-stu-id="0c407-553">Attributes</span></span>| <span data-ttu-id="0c407-554">説明</span><span class="sxs-lookup"><span data-stu-id="0c407-554">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="0c407-555">function</span><span class="sxs-lookup"><span data-stu-id="0c407-555">function</span></span>||<span data-ttu-id="0c407-556">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="0c407-556">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="0c407-557">トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="0c407-557">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="0c407-558">Object</span><span class="sxs-lookup"><span data-stu-id="0c407-558">Object</span></span>| <span data-ttu-id="0c407-559">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="0c407-559">&lt;optional&gt;</span></span>|<span data-ttu-id="0c407-560">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="0c407-560">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0c407-561">要件</span><span class="sxs-lookup"><span data-stu-id="0c407-561">Requirements</span></span>

|<span data-ttu-id="0c407-562">要件</span><span class="sxs-lookup"><span data-stu-id="0c407-562">Requirement</span></span>| <span data-ttu-id="0c407-563">値</span><span class="sxs-lookup"><span data-stu-id="0c407-563">Value</span></span>|
|---|---|
|[<span data-ttu-id="0c407-564">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0c407-564">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0c407-565">1.0</span><span class="sxs-lookup"><span data-stu-id="0c407-565">1.0</span></span>|
|[<span data-ttu-id="0c407-566">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0c407-566">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0c407-567">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0c407-567">ReadItem</span></span>|
|[<span data-ttu-id="0c407-568">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0c407-568">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0c407-569">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="0c407-569">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="0c407-570">例</span><span class="sxs-lookup"><span data-stu-id="0c407-570">Example</span></span>

```js
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="0c407-571">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="0c407-571">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="0c407-572">ユーザーのメールボックスをホストしている Exchange サーバー上の Exchange Web サービス (EWS) のサービスに対して非同期の要求を行います。</span><span class="sxs-lookup"><span data-stu-id="0c407-572">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="0c407-573">このメソッドは、次のシナリオではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="0c407-573">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="0c407-574">Outlook for iOS または Outlook for Android を使用している場合</span><span class="sxs-lookup"><span data-stu-id="0c407-574">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="0c407-575">アドインが Gmail のメールボックスに読み込まれる場合</span><span class="sxs-lookup"><span data-stu-id="0c407-575">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="0c407-576">このような場合は、アドインでは [REST API を使用](https://docs.microsoft.com/outlook/add-ins/use-rest-api)して、代わりにユーザーのメールボックスにアクセスする必要があります。</span><span class="sxs-lookup"><span data-stu-id="0c407-576">In these cases, add-ins should [use REST APIs](https://docs.microsoft.com/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="0c407-577">`makeEwsRequestAsync` メソッドは、アドインの代わりに Exchange に EWS 要求を送信します。</span><span class="sxs-lookup"><span data-stu-id="0c407-577">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="0c407-578">サポートされている EWS 操作の一覧については、「[Outlook アドインからの Web サービスの呼び出し](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="0c407-578">See [Call web services from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="0c407-579">`makeEwsRequestAsync` メソッドでは、フォルダー関連アイテムを要求できません。</span><span class="sxs-lookup"><span data-stu-id="0c407-579">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="0c407-580">XML 要求では UTF-8 エンコードを指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="0c407-580">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="0c407-p143">`makeEwsRequestAsync` メソッドを使用するには、アドインに **ReadWriteMailbox** アクセス許可が必要です。**ReadWriteMailbox** アクセス許可と、`makeEwsRequestAsync` メソッドで呼び出せる EWS 操作の使い方については、「[ユーザーのメールボックスへのメール アドイン アクセスのアクセス許可を指定する](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="0c407-p143">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="0c407-583">サーバー管理者は、クライアント アクセス サーバーの EWS ディレクトリで `OAuthAuthentication` を true に設定して、`makeEwsRequestAsync` メソッドで EWS 要求を行うことができるようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="0c407-583">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="0c407-584">バージョンの相違点</span><span class="sxs-lookup"><span data-stu-id="0c407-584">Version differences</span></span>

<span data-ttu-id="0c407-585">バージョン 15.0.4535.1004 より前のバージョンの Outlook で実行しているメール アプリで `makeEwsRequestAsync` メソッドを使う場合は、エンコード値を `ISO-8859-1` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="0c407-585">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="0c407-p144">Outlook on the web でメール アプリを実行している場合は、エンコード値を設定する必要はありません。mailbox.diagnostics.hostName プロパティを使って、メール アプリを Outlook で実行しているのか、Outlook on the web で実行しているのかを確認できます。mailbox.diagnostics.hostVersion プロパティを使って、どのバージョンの Outlook を使って実行しているかを確認できます。</span><span class="sxs-lookup"><span data-stu-id="0c407-p144">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0c407-589">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="0c407-589">Parameters:</span></span>

|<span data-ttu-id="0c407-590">名前</span><span class="sxs-lookup"><span data-stu-id="0c407-590">Name</span></span>| <span data-ttu-id="0c407-591">型</span><span class="sxs-lookup"><span data-stu-id="0c407-591">Type</span></span>| <span data-ttu-id="0c407-592">属性</span><span class="sxs-lookup"><span data-stu-id="0c407-592">Attributes</span></span>| <span data-ttu-id="0c407-593">説明</span><span class="sxs-lookup"><span data-stu-id="0c407-593">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="0c407-594">String</span><span class="sxs-lookup"><span data-stu-id="0c407-594">String</span></span>||<span data-ttu-id="0c407-595">EWS 要求です。</span><span class="sxs-lookup"><span data-stu-id="0c407-595">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="0c407-596">function</span><span class="sxs-lookup"><span data-stu-id="0c407-596">function</span></span>||<span data-ttu-id="0c407-597">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="0c407-597">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="0c407-598">EWS 呼び出しの XML 結果は、`asyncResult.value` プロパティ内の文字列として提供されています。</span><span class="sxs-lookup"><span data-stu-id="0c407-598">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="0c407-599">結果のサイズが 1 MB を超える場合、代わりにエラー メッセージが返されます。</span><span class="sxs-lookup"><span data-stu-id="0c407-599">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="0c407-600">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="0c407-600">Object</span></span>| <span data-ttu-id="0c407-601">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="0c407-601">&lt;optional&gt;</span></span>|<span data-ttu-id="0c407-602">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="0c407-602">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0c407-603">要件</span><span class="sxs-lookup"><span data-stu-id="0c407-603">Requirements</span></span>

|<span data-ttu-id="0c407-604">要件</span><span class="sxs-lookup"><span data-stu-id="0c407-604">Requirement</span></span>| <span data-ttu-id="0c407-605">値</span><span class="sxs-lookup"><span data-stu-id="0c407-605">Value</span></span>|
|---|---|
|[<span data-ttu-id="0c407-606">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0c407-606">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0c407-607">1.0</span><span class="sxs-lookup"><span data-stu-id="0c407-607">1.0</span></span>|
|[<span data-ttu-id="0c407-608">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0c407-608">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0c407-609">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="0c407-609">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="0c407-610">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0c407-610">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0c407-611">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="0c407-611">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="0c407-612">例</span><span class="sxs-lookup"><span data-stu-id="0c407-612">Example</span></span>

<span data-ttu-id="0c407-613">次の例は、`GetItem` 操作を使ってアイテムの件名を取得するため、`makeEwsRequestAsync` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="0c407-613">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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

####  <a name="removehandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="0c407-614">removeHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="0c407-614">removeHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="0c407-615">サポートされているイベントのイベント ハンドラーを削除します。</span><span class="sxs-lookup"><span data-stu-id="0c407-615">Removes an event handler for a supported event.</span></span>

<span data-ttu-id="0c407-616">現在、サポートされているイベントの種類は `Office.EventType.ItemChanged` だけです。</span><span class="sxs-lookup"><span data-stu-id="0c407-616">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0c407-617">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="0c407-617">Parameters:</span></span>

| <span data-ttu-id="0c407-618">名前</span><span class="sxs-lookup"><span data-stu-id="0c407-618">Name</span></span> | <span data-ttu-id="0c407-619">型</span><span class="sxs-lookup"><span data-stu-id="0c407-619">Type</span></span> | <span data-ttu-id="0c407-620">属性</span><span class="sxs-lookup"><span data-stu-id="0c407-620">Attributes</span></span> | <span data-ttu-id="0c407-621">説明</span><span class="sxs-lookup"><span data-stu-id="0c407-621">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="0c407-622">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="0c407-622">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="0c407-623">ハンドラーを取り消すイベント。</span><span class="sxs-lookup"><span data-stu-id="0c407-623">The event that should revoke the handler.</span></span> |
| `handler` | <span data-ttu-id="0c407-624">職務</span><span class="sxs-lookup"><span data-stu-id="0c407-624">Function</span></span> || <span data-ttu-id="0c407-p146">イベントを処理する関数。関数は、オブジェクト リテラルである単一パラメーターを受け入れる必要があります。パラメーターの `type` プロパティは、`addHandlerAsync` に渡される `eventType` パラメーターと一致します。</span><span class="sxs-lookup"><span data-stu-id="0c407-p146">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="0c407-628">Object</span><span class="sxs-lookup"><span data-stu-id="0c407-628">Object</span></span> | <span data-ttu-id="0c407-629">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0c407-629">&lt;optional&gt;</span></span> | <span data-ttu-id="0c407-630">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="0c407-630">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="0c407-631">Object</span><span class="sxs-lookup"><span data-stu-id="0c407-631">Object</span></span> | <span data-ttu-id="0c407-632">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0c407-632">&lt;optional&gt;</span></span> | <span data-ttu-id="0c407-633">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="0c407-633">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="0c407-634">function</span><span class="sxs-lookup"><span data-stu-id="0c407-634">function</span></span>| <span data-ttu-id="0c407-635">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0c407-635">&lt;optional&gt;</span></span>|<span data-ttu-id="0c407-636">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="0c407-636">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0c407-637">要件</span><span class="sxs-lookup"><span data-stu-id="0c407-637">Requirements</span></span>

|<span data-ttu-id="0c407-638">要件</span><span class="sxs-lookup"><span data-stu-id="0c407-638">Requirement</span></span>| <span data-ttu-id="0c407-639">値</span><span class="sxs-lookup"><span data-stu-id="0c407-639">Value</span></span>|
|---|---|
|[<span data-ttu-id="0c407-640">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0c407-640">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0c407-641">1.5</span><span class="sxs-lookup"><span data-stu-id="0c407-641">1.5</span></span> |
|[<span data-ttu-id="0c407-642">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0c407-642">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0c407-643">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0c407-643">ReadItem</span></span> |
|[<span data-ttu-id="0c407-644">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0c407-644">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0c407-645">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="0c407-645">Compose or read</span></span>|