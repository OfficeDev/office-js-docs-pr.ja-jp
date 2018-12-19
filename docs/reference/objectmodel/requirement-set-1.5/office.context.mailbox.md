
# <a name="mailbox"></a><span data-ttu-id="a5143-101">mailbox</span><span class="sxs-lookup"><span data-stu-id="a5143-101">mailbox</span></span>

### <span data-ttu-id="a5143-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span><span class="sxs-lookup"><span data-stu-id="a5143-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span></span>

<span data-ttu-id="a5143-104">Microsoft Outlook と Microsoft Outlook on the web の Outlook アドイン オブジェクト モデルへのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="a5143-104">Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a5143-105">要件</span><span class="sxs-lookup"><span data-stu-id="a5143-105">Requirements</span></span>

|<span data-ttu-id="a5143-106">要件</span><span class="sxs-lookup"><span data-stu-id="a5143-106">Requirement</span></span>| <span data-ttu-id="a5143-107">値</span><span class="sxs-lookup"><span data-stu-id="a5143-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="a5143-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a5143-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a5143-109">1.0</span><span class="sxs-lookup"><span data-stu-id="a5143-109">1.0</span></span>|
|[<span data-ttu-id="a5143-110">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="a5143-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a5143-111">制限あり</span><span class="sxs-lookup"><span data-stu-id="a5143-111">Restricted</span></span>|
|[<span data-ttu-id="a5143-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a5143-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a5143-113">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="a5143-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="a5143-114">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="a5143-114">Members and methods</span></span>

| <span data-ttu-id="a5143-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="a5143-115">Member</span></span> | <span data-ttu-id="a5143-116">種類</span><span class="sxs-lookup"><span data-stu-id="a5143-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="a5143-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="a5143-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="a5143-118">メンバー</span><span class="sxs-lookup"><span data-stu-id="a5143-118">Member</span></span> |
| [<span data-ttu-id="a5143-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="a5143-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="a5143-120">メンバー</span><span class="sxs-lookup"><span data-stu-id="a5143-120">Member</span></span> |
| [<span data-ttu-id="a5143-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="a5143-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="a5143-122">メソッド</span><span class="sxs-lookup"><span data-stu-id="a5143-122">Method</span></span> |
| [<span data-ttu-id="a5143-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="a5143-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="a5143-124">メソッド</span><span class="sxs-lookup"><span data-stu-id="a5143-124">Method</span></span> |
| [<span data-ttu-id="a5143-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="a5143-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) | <span data-ttu-id="a5143-126">メソッド</span><span class="sxs-lookup"><span data-stu-id="a5143-126">Method</span></span> |
| [<span data-ttu-id="a5143-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="a5143-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="a5143-128">メソッド</span><span class="sxs-lookup"><span data-stu-id="a5143-128">Method</span></span> |
| [<span data-ttu-id="a5143-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="a5143-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="a5143-130">メソッド</span><span class="sxs-lookup"><span data-stu-id="a5143-130">Method</span></span> |
| [<span data-ttu-id="a5143-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="a5143-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="a5143-132">メソッド</span><span class="sxs-lookup"><span data-stu-id="a5143-132">Method</span></span> |
| [<span data-ttu-id="a5143-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="a5143-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="a5143-134">メソッド</span><span class="sxs-lookup"><span data-stu-id="a5143-134">Method</span></span> |
| [<span data-ttu-id="a5143-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="a5143-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="a5143-136">メソッド</span><span class="sxs-lookup"><span data-stu-id="a5143-136">Method</span></span> |
| [<span data-ttu-id="a5143-137">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="a5143-137">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="a5143-138">メソッド</span><span class="sxs-lookup"><span data-stu-id="a5143-138">Method</span></span> |
| [<span data-ttu-id="a5143-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="a5143-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="a5143-140">メソッド</span><span class="sxs-lookup"><span data-stu-id="a5143-140">Method</span></span> |
| [<span data-ttu-id="a5143-141">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="a5143-141">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="a5143-142">メソッド</span><span class="sxs-lookup"><span data-stu-id="a5143-142">Method</span></span> |
| [<span data-ttu-id="a5143-143">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="a5143-143">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="a5143-144">メソッド</span><span class="sxs-lookup"><span data-stu-id="a5143-144">Method</span></span> |
| [<span data-ttu-id="a5143-145">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="a5143-145">removeHandlerAsync</span></span>](#removehandlerasynceventtype-handler-options-callback) | <span data-ttu-id="a5143-146">メソッド</span><span class="sxs-lookup"><span data-stu-id="a5143-146">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="a5143-147">名前空間</span><span class="sxs-lookup"><span data-stu-id="a5143-147">Namespaces</span></span>

<span data-ttu-id="a5143-148">[diagnostics](Office.context.mailbox.diagnostics.md):Outlook アドインに診断情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="a5143-148">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="a5143-149">[item](Office.context.mailbox.item.md):Outlook アドインのメッセージや予定にアクセスするためのメソッドとプロパティを提供します。</span><span class="sxs-lookup"><span data-stu-id="a5143-149">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="a5143-150">[userProfile](Office.context.mailbox.userProfile.md):Outlook アドインのユーザーに関する情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="a5143-150">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="a5143-151">メンバー</span><span class="sxs-lookup"><span data-stu-id="a5143-151">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="a5143-152">ewsUrl :String</span><span class="sxs-lookup"><span data-stu-id="a5143-152">ewsUrl :String</span></span>

<span data-ttu-id="a5143-p102">このメール アカウントの Exchange Web サービス (EWS) エンドポイントの URL を取得します。読み取りモードのみです。</span><span class="sxs-lookup"><span data-stu-id="a5143-p102">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="a5143-155">このメンバーは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="a5143-155">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="a5143-p103">
  `ewsUrl\` 値は、リモート サービスで、ユーザーのメールボックスに EWS 呼び出しを行うために使うことができます。たとえば、[選択したアイテムから添付ファイルを取得する](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="a5143-p103">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="a5143-158">アプリが閲覧モードで `ewsUrl` メンバーを呼び出すには、アプリのマニフェスト内に **ReadItem** アクセス許可が指定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="a5143-158">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="a5143-p104">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出してから、`ewsUrl` メンバーを使用する必要があります。アプリには、`saveAsync` メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="a5143-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="a5143-161">型:</span><span class="sxs-lookup"><span data-stu-id="a5143-161">Type:</span></span>

*   <span data-ttu-id="a5143-162">String</span><span class="sxs-lookup"><span data-stu-id="a5143-162">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a5143-163">要件</span><span class="sxs-lookup"><span data-stu-id="a5143-163">Requirements</span></span>

|<span data-ttu-id="a5143-164">要件</span><span class="sxs-lookup"><span data-stu-id="a5143-164">Requirement</span></span>| <span data-ttu-id="a5143-165">値</span><span class="sxs-lookup"><span data-stu-id="a5143-165">Value</span></span>|
|---|---|
|[<span data-ttu-id="a5143-166">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a5143-166">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a5143-167">1.0</span><span class="sxs-lookup"><span data-stu-id="a5143-167">1.0</span></span>|
|[<span data-ttu-id="a5143-168">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="a5143-168">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a5143-169">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a5143-169">ReadItem</span></span>|
|[<span data-ttu-id="a5143-170">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a5143-170">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a5143-171">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="a5143-171">Compose or read</span></span>|

#### <a name="resturl-string"></a><span data-ttu-id="a5143-172">restUrl :String</span><span class="sxs-lookup"><span data-stu-id="a5143-172">restUrl :String</span></span>

<span data-ttu-id="a5143-173">この電子メール アカウントの REST エンドポイントの URL を取得します。</span><span class="sxs-lookup"><span data-stu-id="a5143-173">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="a5143-174">`restUrl` 値は、ユーザーのメールボックスに [REST API](https://docs.microsoft.com/outlook/rest/) 呼び出しを行うために使用できます。</span><span class="sxs-lookup"><span data-stu-id="a5143-174">The `restUrl` value can be used to make [REST API](https://docs.microsoft.com/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="a5143-175">アプリが閲覧モードで `restUrl` メンバーを呼び出すには、アプリのマニフェスト内に **ReadItem** アクセス許可が指定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="a5143-175">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="a5143-p105">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出してから、`restUrl` メンバーを使用する必要があります。アプリには、`saveAsync` メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="a5143-p105">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

> [!NOTE]
> <span data-ttu-id="a5143-178">構成されたカスタム REST URL を使用する Exchange 2016 以降のオンプレミスのインストールに接続されている Outlook クライアントは、`restUrl` に無効な値を返します。</span><span class="sxs-lookup"><span data-stu-id="a5143-178">Note: Outlook clients connected to on-premises installations of Exchange 2016 with a custom REST URL configured will return an invalid value for `restUrl`.</span></span>

##### <a name="type"></a><span data-ttu-id="a5143-179">型:</span><span class="sxs-lookup"><span data-stu-id="a5143-179">Type:</span></span>

*   <span data-ttu-id="a5143-180">String</span><span class="sxs-lookup"><span data-stu-id="a5143-180">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a5143-181">要件</span><span class="sxs-lookup"><span data-stu-id="a5143-181">Requirements</span></span>

|<span data-ttu-id="a5143-182">要件</span><span class="sxs-lookup"><span data-stu-id="a5143-182">Requirement</span></span>| <span data-ttu-id="a5143-183">値</span><span class="sxs-lookup"><span data-stu-id="a5143-183">Value</span></span>|
|---|---|
|[<span data-ttu-id="a5143-184">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a5143-184">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a5143-185">1.5</span><span class="sxs-lookup"><span data-stu-id="a5143-185">1.5</span></span> |
|[<span data-ttu-id="a5143-186">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="a5143-186">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a5143-187">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a5143-187">ReadItem</span></span>|
|[<span data-ttu-id="a5143-188">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a5143-188">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a5143-189">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="a5143-189">Compose or read</span></span>|

### <a name="methods"></a><span data-ttu-id="a5143-190">メソッド</span><span class="sxs-lookup"><span data-stu-id="a5143-190">Methods</span></span>

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="a5143-191">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="a5143-191">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="a5143-192">サポートされているイベントのイベント ハンドラーを追加します。</span><span class="sxs-lookup"><span data-stu-id="a5143-192">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="a5143-193">現在サポートされている唯一のイベントの種類は `Office.EventType.ItemChanged` です。これはユーザーが新しいアイテムを選択したときに呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="a5143-193">Currently the only supported event type is `Office.EventType.ItemChanged`, which is invoked when the user selects a new item.</span></span> <span data-ttu-id="a5143-194">このイベントは、ピン留め可能な作業ウィンドウを実装するアドインで使用され、現在選択されているアイテムに基づいて作業ウィンドウ UI をアドインで更新できるようにします。</span><span class="sxs-lookup"><span data-stu-id="a5143-194">Currently the only supported event type is , which is invoked when the user selects a new item. This event is used by add-ins that implement a pinnable taskpane, and allows the add-in to refresh the taskpane UI based on the currently selected item.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a5143-195">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="a5143-195">Parameters:</span></span>

| <span data-ttu-id="a5143-196">名前</span><span class="sxs-lookup"><span data-stu-id="a5143-196">Name</span></span> | <span data-ttu-id="a5143-197">型</span><span class="sxs-lookup"><span data-stu-id="a5143-197">Type</span></span> | <span data-ttu-id="a5143-198">属性</span><span class="sxs-lookup"><span data-stu-id="a5143-198">Attributes</span></span> | <span data-ttu-id="a5143-199">説明</span><span class="sxs-lookup"><span data-stu-id="a5143-199">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="a5143-200">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="a5143-200">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="a5143-201">ハンドラーを呼び出す必要のあるイベント。</span><span class="sxs-lookup"><span data-stu-id="a5143-201">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="a5143-202">Function</span><span class="sxs-lookup"><span data-stu-id="a5143-202">Function</span></span> || <span data-ttu-id="a5143-p107">イベントを処理する関数。関数は、オブジェクト リテラルである単一パラメーターを受け入れる必要があります。パラメーターの `type` プロパティは、`addHandlerAsync` に渡される `eventType` パラメーターと一致します。</span><span class="sxs-lookup"><span data-stu-id="a5143-p107">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="a5143-206">Object</span><span class="sxs-lookup"><span data-stu-id="a5143-206">Object</span></span> | <span data-ttu-id="a5143-207">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a5143-207">&lt;optional&gt;</span></span> | <span data-ttu-id="a5143-208">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="a5143-208">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="a5143-209">Object</span><span class="sxs-lookup"><span data-stu-id="a5143-209">Object</span></span> | <span data-ttu-id="a5143-210">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a5143-210">&lt;optional&gt;</span></span> | <span data-ttu-id="a5143-211">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="a5143-211">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="a5143-212">function</span><span class="sxs-lookup"><span data-stu-id="a5143-212">function</span></span>| <span data-ttu-id="a5143-213">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a5143-213">&lt;optional&gt;</span></span>|<span data-ttu-id="a5143-214">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="a5143-214">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a5143-215">要件</span><span class="sxs-lookup"><span data-stu-id="a5143-215">Requirements</span></span>

|<span data-ttu-id="a5143-216">要件</span><span class="sxs-lookup"><span data-stu-id="a5143-216">Requirement</span></span>| <span data-ttu-id="a5143-217">値</span><span class="sxs-lookup"><span data-stu-id="a5143-217">Value</span></span>|
|---|---|
|[<span data-ttu-id="a5143-218">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a5143-218">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a5143-219">1.5</span><span class="sxs-lookup"><span data-stu-id="a5143-219">1.5</span></span> |
|[<span data-ttu-id="a5143-220">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="a5143-220">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a5143-221">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a5143-221">ReadItem</span></span> |
|[<span data-ttu-id="a5143-222">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a5143-222">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a5143-223">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="a5143-223">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a5143-224">例</span><span class="sxs-lookup"><span data-stu-id="a5143-224">Example</span></span>

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

####  <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="a5143-225">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="a5143-225">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="a5143-226">REST 形式のアイテム ID を EWS 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="a5143-226">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="a5143-227">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="a5143-227">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="a5143-p108">REST API ([Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) や [Microsoft Graph](https://graph.microsoft.io/) など) で取得されたアイテム ID は、Exchange Web サービス (EWS) に使用される形式とは異なる形式を使用します。`convertToEwsId` メソッドは、REST 形式の ID を EWS 用の適切な形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="a5143-p108">Item IDs retrieved via a REST API (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a5143-230">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="a5143-230">Parameters:</span></span>

|<span data-ttu-id="a5143-231">名前</span><span class="sxs-lookup"><span data-stu-id="a5143-231">Name</span></span>| <span data-ttu-id="a5143-232">型</span><span class="sxs-lookup"><span data-stu-id="a5143-232">Type</span></span>| <span data-ttu-id="a5143-233">説明</span><span class="sxs-lookup"><span data-stu-id="a5143-233">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="a5143-234">String</span><span class="sxs-lookup"><span data-stu-id="a5143-234">String</span></span>|<span data-ttu-id="a5143-235">Outlook REST API 形式のアイテム ID</span><span class="sxs-lookup"><span data-stu-id="a5143-235">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="a5143-236">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="a5143-236">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.restversion)|<span data-ttu-id="a5143-237">アイテム ID の取得に使用された Outlook REST API のバージョンを示す値。</span><span class="sxs-lookup"><span data-stu-id="a5143-237">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a5143-238">要件</span><span class="sxs-lookup"><span data-stu-id="a5143-238">Requirements</span></span>

|<span data-ttu-id="a5143-239">要件</span><span class="sxs-lookup"><span data-stu-id="a5143-239">Requirement</span></span>| <span data-ttu-id="a5143-240">値</span><span class="sxs-lookup"><span data-stu-id="a5143-240">Value</span></span>|
|---|---|
|[<span data-ttu-id="a5143-241">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a5143-241">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a5143-242">1.3</span><span class="sxs-lookup"><span data-stu-id="a5143-242">1.3</span></span>|
|[<span data-ttu-id="a5143-243">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="a5143-243">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a5143-244">制限あり</span><span class="sxs-lookup"><span data-stu-id="a5143-244">Restricted</span></span>|
|[<span data-ttu-id="a5143-245">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a5143-245">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a5143-246">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="a5143-246">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a5143-247">戻り値:</span><span class="sxs-lookup"><span data-stu-id="a5143-247">Returns:</span></span>

<span data-ttu-id="a5143-248">型:String</span><span class="sxs-lookup"><span data-stu-id="a5143-248">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="a5143-249">例</span><span class="sxs-lookup"><span data-stu-id="a5143-249">Example</span></span>

```js
// Get an item's ID from a REST API
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the
// Outlook Mail API
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime"></a><span data-ttu-id="a5143-250">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_5/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="a5143-250">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_5/office.LocalClientTime)}</span></span>

<span data-ttu-id="a5143-251">クライアントのローカル時間で時間情報が含まれている辞書を取得します。</span><span class="sxs-lookup"><span data-stu-id="a5143-251">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="a5143-p109">Outlook 用メール アプリや Outlook Web App で使う日付と時刻では、異なるタイム ゾーンを使うことができます。Outlook では、クライアント コンピューターのタイム ゾーンを使います。Outlook Web App では、Exchange 管理センター (EAC) で設定されたタイム ゾーンを使います。ユーザー インターフェイスに表示される値が、常にユーザーが期待するタイム ゾーンと一致するように日付と時刻の値を処理する必要があります。</span><span class="sxs-lookup"><span data-stu-id="a5143-p109">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="a5143-p110">Outlook でメール アプリが実行されている場合、`convertToLocalClientTime` メソッドは、クライアント コンピューターのタイム ゾーンに設定された値のディクショナリ オブジェクトを返します。Outlook Web Apps でメール アプリが実行されている場合、`convertToLocalClientTime` メソッドは、EAC に指定されたタイム ゾーンに設定された値のディクショナリ オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="a5143-p110">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a5143-257">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="a5143-257">Parameters:</span></span>

|<span data-ttu-id="a5143-258">名前</span><span class="sxs-lookup"><span data-stu-id="a5143-258">Name</span></span>| <span data-ttu-id="a5143-259">型</span><span class="sxs-lookup"><span data-stu-id="a5143-259">Type</span></span>| <span data-ttu-id="a5143-260">説明</span><span class="sxs-lookup"><span data-stu-id="a5143-260">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="a5143-261">Date</span><span class="sxs-lookup"><span data-stu-id="a5143-261">Date</span></span>|<span data-ttu-id="a5143-262">日付オブジェクト</span><span class="sxs-lookup"><span data-stu-id="a5143-262">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a5143-263">要件</span><span class="sxs-lookup"><span data-stu-id="a5143-263">Requirements</span></span>

|<span data-ttu-id="a5143-264">要件</span><span class="sxs-lookup"><span data-stu-id="a5143-264">Requirement</span></span>| <span data-ttu-id="a5143-265">値</span><span class="sxs-lookup"><span data-stu-id="a5143-265">Value</span></span>|
|---|---|
|[<span data-ttu-id="a5143-266">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a5143-266">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a5143-267">1.0</span><span class="sxs-lookup"><span data-stu-id="a5143-267">1.0</span></span>|
|[<span data-ttu-id="a5143-268">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="a5143-268">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a5143-269">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a5143-269">ReadItem</span></span>|
|[<span data-ttu-id="a5143-270">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a5143-270">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a5143-271">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="a5143-271">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a5143-272">戻り値:</span><span class="sxs-lookup"><span data-stu-id="a5143-272">Returns:</span></span>

<span data-ttu-id="a5143-273">型:[LocalClientTime](/javascript/api/outlook_1_5/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="a5143-273">Type: [LocalClientTime](/javascript/api/outlook_1_5/office.LocalClientTime)</span></span>

####  <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="a5143-274">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="a5143-274">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="a5143-275">EWS 形式のアイテム ID を REST 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="a5143-275">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="a5143-276">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="a5143-276">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="a5143-p111">EWS または `itemId` プロパティで取得されるアイテム ID は、REST API ([Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) や [Microsoft Graph](https://graph.microsoft.io/) など) に使用される形式とは異なる形式を使用します。`convertToRestId` メソッドは、EWS 形式の ID を REST 用の適切な形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="a5143-p111">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a5143-279">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="a5143-279">Parameters:</span></span>

|<span data-ttu-id="a5143-280">名前</span><span class="sxs-lookup"><span data-stu-id="a5143-280">Name</span></span>| <span data-ttu-id="a5143-281">型</span><span class="sxs-lookup"><span data-stu-id="a5143-281">Type</span></span>| <span data-ttu-id="a5143-282">説明</span><span class="sxs-lookup"><span data-stu-id="a5143-282">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="a5143-283">String</span><span class="sxs-lookup"><span data-stu-id="a5143-283">String</span></span>|<span data-ttu-id="a5143-284">Exchange Web サービス (EWS) 形式のアイテム ID</span><span class="sxs-lookup"><span data-stu-id="a5143-284">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="a5143-285">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="a5143-285">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.restversion)|<span data-ttu-id="a5143-286">変換後の ID を使用する Outlook REST API のバージョンを示す値。</span><span class="sxs-lookup"><span data-stu-id="a5143-286">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a5143-287">要件</span><span class="sxs-lookup"><span data-stu-id="a5143-287">Requirements</span></span>

|<span data-ttu-id="a5143-288">要件</span><span class="sxs-lookup"><span data-stu-id="a5143-288">Requirement</span></span>| <span data-ttu-id="a5143-289">値</span><span class="sxs-lookup"><span data-stu-id="a5143-289">Value</span></span>|
|---|---|
|[<span data-ttu-id="a5143-290">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a5143-290">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a5143-291">1.3</span><span class="sxs-lookup"><span data-stu-id="a5143-291">1.3</span></span>|
|[<span data-ttu-id="a5143-292">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="a5143-292">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a5143-293">制限あり</span><span class="sxs-lookup"><span data-stu-id="a5143-293">Restricted</span></span>|
|[<span data-ttu-id="a5143-294">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a5143-294">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a5143-295">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="a5143-295">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a5143-296">戻り値:</span><span class="sxs-lookup"><span data-stu-id="a5143-296">Returns:</span></span>

<span data-ttu-id="a5143-297">型:String</span><span class="sxs-lookup"><span data-stu-id="a5143-297">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="a5143-298">例</span><span class="sxs-lookup"><span data-stu-id="a5143-298">Example</span></span>

```js
// Get the currently selected item's ID
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the
// Outlook Mail API
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="a5143-299">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="a5143-299">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="a5143-300">時間情報が含まれているディクショナリから日付オブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="a5143-300">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="a5143-301">`convertToUtcClientTime` メソッドは、ローカルの日付と時刻を含むディクショナリを、ローカルの日付と時刻の正しい値を持つ日付オブジェクトに変換します。</span><span class="sxs-lookup"><span data-stu-id="a5143-301">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a5143-302">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="a5143-302">Parameters:</span></span>

|<span data-ttu-id="a5143-303">名前</span><span class="sxs-lookup"><span data-stu-id="a5143-303">Name</span></span>| <span data-ttu-id="a5143-304">型</span><span class="sxs-lookup"><span data-stu-id="a5143-304">Type</span></span>| <span data-ttu-id="a5143-305">説明</span><span class="sxs-lookup"><span data-stu-id="a5143-305">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="a5143-306">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="a5143-306">LocalClientTime</span></span>](/javascript/api/outlook_1_5/office.LocalClientTime)|<span data-ttu-id="a5143-307">変換するローカル時刻の値。</span><span class="sxs-lookup"><span data-stu-id="a5143-307">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a5143-308">要件</span><span class="sxs-lookup"><span data-stu-id="a5143-308">Requirements</span></span>

|<span data-ttu-id="a5143-309">要件</span><span class="sxs-lookup"><span data-stu-id="a5143-309">Requirement</span></span>| <span data-ttu-id="a5143-310">値</span><span class="sxs-lookup"><span data-stu-id="a5143-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="a5143-311">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a5143-311">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a5143-312">1.0</span><span class="sxs-lookup"><span data-stu-id="a5143-312">1.0</span></span>|
|[<span data-ttu-id="a5143-313">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="a5143-313">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a5143-314">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a5143-314">ReadItem</span></span>|
|[<span data-ttu-id="a5143-315">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a5143-315">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a5143-316">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="a5143-316">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a5143-317">戻り値:</span><span class="sxs-lookup"><span data-stu-id="a5143-317">Returns:</span></span>

<span data-ttu-id="a5143-318">時間が UTC で表現された日付オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="a5143-318">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="a5143-319">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="a5143-319">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="a5143-320">Date</span><span class="sxs-lookup"><span data-stu-id="a5143-320">Date</span></span></dd>

</dl>

####  <a name="displayappointmentformitemid"></a><span data-ttu-id="a5143-321">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="a5143-321">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="a5143-322">既存の予定を表示します。</span><span class="sxs-lookup"><span data-stu-id="a5143-322">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="a5143-323">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="a5143-323">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="a5143-324">`displayAppointmentForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存の予定を開きます。</span><span class="sxs-lookup"><span data-stu-id="a5143-324">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="a5143-p112">Outlook for Mac では、この方法を使って、定期的な系列の一部ではない単一の予定や定期的な系列のマスター予定を表示できます。ただし、系列のインスタンスは表示できません。これは、Outlook for Mac においては定期的な系列のインスタンスのプロパティ (アイテム ID を含む) にアクセスできないためです。</span><span class="sxs-lookup"><span data-stu-id="a5143-p112">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="a5143-327">Outlook Web App では、このメソッドは指定されたフォームの本文が 32 KB 以下の文字数の場合にフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="a5143-327">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="a5143-328">指定のアイテム識別子が既存の予定を表していない場合は、クライアント コンピューターまたはデバイスで空のウィンドウが開き、エラー メッセージは返されません。</span><span class="sxs-lookup"><span data-stu-id="a5143-328">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a5143-329">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="a5143-329">Parameters:</span></span>

|<span data-ttu-id="a5143-330">名前</span><span class="sxs-lookup"><span data-stu-id="a5143-330">Name</span></span>| <span data-ttu-id="a5143-331">型</span><span class="sxs-lookup"><span data-stu-id="a5143-331">Type</span></span>| <span data-ttu-id="a5143-332">説明</span><span class="sxs-lookup"><span data-stu-id="a5143-332">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="a5143-333">String</span><span class="sxs-lookup"><span data-stu-id="a5143-333">String</span></span>|<span data-ttu-id="a5143-334">既存の予定の Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="a5143-334">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a5143-335">要件</span><span class="sxs-lookup"><span data-stu-id="a5143-335">Requirements</span></span>

|<span data-ttu-id="a5143-336">要件</span><span class="sxs-lookup"><span data-stu-id="a5143-336">Requirement</span></span>| <span data-ttu-id="a5143-337">値</span><span class="sxs-lookup"><span data-stu-id="a5143-337">Value</span></span>|
|---|---|
|[<span data-ttu-id="a5143-338">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a5143-338">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a5143-339">1.0</span><span class="sxs-lookup"><span data-stu-id="a5143-339">1.0</span></span>|
|[<span data-ttu-id="a5143-340">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="a5143-340">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a5143-341">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a5143-341">ReadItem</span></span>|
|[<span data-ttu-id="a5143-342">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a5143-342">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a5143-343">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="a5143-343">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a5143-344">例</span><span class="sxs-lookup"><span data-stu-id="a5143-344">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  <a name="displaymessageformitemid"></a><span data-ttu-id="a5143-345">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="a5143-345">displayMessageForm(itemId)</span></span>

<span data-ttu-id="a5143-346">既存のメッセージを表示します。</span><span class="sxs-lookup"><span data-stu-id="a5143-346">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="a5143-347">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="a5143-347">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="a5143-348">`displayMessageForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存のメッセージを開きます。</span><span class="sxs-lookup"><span data-stu-id="a5143-348">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="a5143-349">Outlook Web App では、このメソッドは指定されたフォームの本文が 32 KB 以下の文字数の場合にフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="a5143-349">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="a5143-350">指定のアイテム識別子が既存のメッセージを表していない場合は、クライアント コンピューターにはメッセージは表示されず、エラー メッセージも返されません。</span><span class="sxs-lookup"><span data-stu-id="a5143-350">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="a5143-p113">予定を表す `itemId` を含む `displayMessageForm` を使用しないでください。既存の予定を表示するには、`displayAppointmentForm` メソッドを使用します。新しい予定を作成するフォームを表示するには、`displayNewAppointmentForm` メソッドを使用します。</span><span class="sxs-lookup"><span data-stu-id="a5143-p113">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a5143-353">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="a5143-353">Parameters:</span></span>

|<span data-ttu-id="a5143-354">名前</span><span class="sxs-lookup"><span data-stu-id="a5143-354">Name</span></span>| <span data-ttu-id="a5143-355">型</span><span class="sxs-lookup"><span data-stu-id="a5143-355">Type</span></span>| <span data-ttu-id="a5143-356">説明</span><span class="sxs-lookup"><span data-stu-id="a5143-356">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="a5143-357">String</span><span class="sxs-lookup"><span data-stu-id="a5143-357">String</span></span>|<span data-ttu-id="a5143-358">既存のメッセージの Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="a5143-358">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a5143-359">要件</span><span class="sxs-lookup"><span data-stu-id="a5143-359">Requirements</span></span>

|<span data-ttu-id="a5143-360">要件</span><span class="sxs-lookup"><span data-stu-id="a5143-360">Requirement</span></span>| <span data-ttu-id="a5143-361">値</span><span class="sxs-lookup"><span data-stu-id="a5143-361">Value</span></span>|
|---|---|
|[<span data-ttu-id="a5143-362">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a5143-362">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a5143-363">1.0</span><span class="sxs-lookup"><span data-stu-id="a5143-363">1.0</span></span>|
|[<span data-ttu-id="a5143-364">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="a5143-364">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a5143-365">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a5143-365">ReadItem</span></span>|
|[<span data-ttu-id="a5143-366">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a5143-366">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a5143-367">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="a5143-367">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a5143-368">例</span><span class="sxs-lookup"><span data-stu-id="a5143-368">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="a5143-369">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="a5143-369">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="a5143-370">新しい予定を作成するためのフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="a5143-370">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="a5143-371">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="a5143-371">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="a5143-p114">`displayNewAppointmentForm` メソッドを使用すると、ユーザーが新しい予定または会議を作成できるフォームが開きます。パラメーターを指定すると、予定のフォーム フィールドにパラメーターの内容が自動的に設定されます。</span><span class="sxs-lookup"><span data-stu-id="a5143-p114">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="a5143-p115">このメソッドは、Outlook Web App と OWA for Devices において、出席者フィールドが含まれるフォームを必ず表示します。入力引数として出席者を指定しないと、このメソッドにより **[保存]** ボタンのあるフォームが表示されます。出席者を指定した場合には、フォームにその出席者と **[送信]** ボタンが表示されます。</span><span class="sxs-lookup"><span data-stu-id="a5143-p115">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="a5143-p116">Outlook リッチ クライアントと Outlook RT で、`requiredAttendees`、`optionalAttendees`、または `resources` パラメーターに出席者またはリソースを指定し、このメソッドを実行すると、**[送信]** ボタンがある会議フォームが表示されます。受信者を指定せずにこのメソッドを実行すると、**[保存して閉じる]** ボタンがある予定フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="a5143-p116">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="a5143-379">パラメーターのいずれかが指定のサイズ制限を超える場合、または不明なパラメーター名が指定されている場合は、例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="a5143-379">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a5143-380">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="a5143-380">Parameters:</span></span>

|<span data-ttu-id="a5143-381">名前</span><span class="sxs-lookup"><span data-stu-id="a5143-381">Name</span></span>| <span data-ttu-id="a5143-382">型</span><span class="sxs-lookup"><span data-stu-id="a5143-382">Type</span></span>| <span data-ttu-id="a5143-383">説明</span><span class="sxs-lookup"><span data-stu-id="a5143-383">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="a5143-384">Object</span><span class="sxs-lookup"><span data-stu-id="a5143-384">Object</span></span> | <span data-ttu-id="a5143-385">新しい予定を記述するパラメーターのディクショナリ。</span><span class="sxs-lookup"><span data-stu-id="a5143-385">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="a5143-386">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="a5143-386">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="a5143-p117">予定に必要な各出席者について、メール アドレスを含む文字列の配列、または `EmailAddressDetails` オブジェクトを含む配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="a5143-p117">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="a5143-389">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="a5143-389">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="a5143-p118">予定の各任意出席者について、メール アドレスを含む文字列の配列、または `EmailAddressDetails` オブジェクトを含む配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="a5143-p118">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="a5143-392">日付</span><span class="sxs-lookup"><span data-stu-id="a5143-392">Date</span></span> | <span data-ttu-id="a5143-393">予定の開始日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="a5143-393">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="a5143-394">Date</span><span class="sxs-lookup"><span data-stu-id="a5143-394">Date</span></span> | <span data-ttu-id="a5143-395">予定の終了日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="a5143-395">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="a5143-396">String</span><span class="sxs-lookup"><span data-stu-id="a5143-396">String</span></span> | <span data-ttu-id="a5143-p119">予定の場所を含む文字列。文字列は最大 255 文字に制限されます。</span><span class="sxs-lookup"><span data-stu-id="a5143-p119">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="a5143-399">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="a5143-399">Array.&lt;String&gt;</span></span> | <span data-ttu-id="a5143-p120">予定に必要なリソースを含む文字列の配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="a5143-p120">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="a5143-402">String</span><span class="sxs-lookup"><span data-stu-id="a5143-402">String</span></span> | <span data-ttu-id="a5143-p121">予定の件名を含む文字列です。文字列は最大 255 文字に制限されます。</span><span class="sxs-lookup"><span data-stu-id="a5143-p121">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="a5143-405">String</span><span class="sxs-lookup"><span data-stu-id="a5143-405">String</span></span> | <span data-ttu-id="a5143-p122">予定の本文。本文の内容は、最大サイズが 32 KB に制限されます。</span><span class="sxs-lookup"><span data-stu-id="a5143-p122">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="a5143-408">要件</span><span class="sxs-lookup"><span data-stu-id="a5143-408">Requirements</span></span>

|<span data-ttu-id="a5143-409">要件</span><span class="sxs-lookup"><span data-stu-id="a5143-409">Requirement</span></span>| <span data-ttu-id="a5143-410">値</span><span class="sxs-lookup"><span data-stu-id="a5143-410">Value</span></span>|
|---|---|
|[<span data-ttu-id="a5143-411">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a5143-411">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a5143-412">1.0</span><span class="sxs-lookup"><span data-stu-id="a5143-412">1.0</span></span>|
|[<span data-ttu-id="a5143-413">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="a5143-413">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a5143-414">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a5143-414">ReadItem</span></span>|
|[<span data-ttu-id="a5143-415">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a5143-415">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a5143-416">読み取り</span><span class="sxs-lookup"><span data-stu-id="a5143-416">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a5143-417">例</span><span class="sxs-lookup"><span data-stu-id="a5143-417">Example</span></span>

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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="a5143-418">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="a5143-418">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="a5143-419">REST API または Exchange Web サービスを呼び出すために使用するトークンを含む文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="a5143-419">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="a5143-p123">`getCallbackTokenAsync` メソッドは、ユーザーのメールボックスをホストする Exchange Server から不透明なトークンを取得する非同期の呼び出しを行います。コールバック トークンの有効期間は 5 分です。</span><span class="sxs-lookup"><span data-stu-id="a5143-p123">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="a5143-422">可能であれば、アドインでは Exchange Web サービスの代わりに REST API を使用することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="a5143-422">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span> 

<span data-ttu-id="a5143-423">**REST トークン**</span><span class="sxs-lookup"><span data-stu-id="a5143-423">**REST Tokens**</span></span>

<span data-ttu-id="a5143-p124">REST トークンが要求された場合 (`options.isRest = true`)、結果トークンは Exchange Web サービスの呼び出しを認証するためには機能しません。アドインがマニフェストで [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) アクセス許可を指定していない限り、トークンの範囲は現在のアイテムとその添付ファイルへの読み取り専用アクセスに制限されます。`ReadWriteMailbox` アクセス許可が指定されている場合は、結果トークンは、メールを送信する機能など、メール、カレンダー、連絡先への読み取り/書き込みアクセスを付与します。</span><span class="sxs-lookup"><span data-stu-id="a5143-p124">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="a5143-427">アドインでは、`restUrl` プロパティを使用して、REST API 呼び出しを行うときに使用する正しい URL を決定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="a5143-427">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="a5143-428">**EWS トークン**</span><span class="sxs-lookup"><span data-stu-id="a5143-428">**EWS Tokens**</span></span>

<span data-ttu-id="a5143-p125">EWS トークンが要求された場合 (`options.isRest = false`)、結果トークンは REST API 呼び出しを認証するためには機能しません。トークンの範囲は、現在のアイテムへのアクセスに制限されます。</span><span class="sxs-lookup"><span data-stu-id="a5143-p125">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="a5143-431">アドインでは、`ewsUrl` プロパティを使用して、EWS 呼び出しを行うときに使用する正しい URL を決定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="a5143-431">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a5143-432">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="a5143-432">Parameters:</span></span>

|<span data-ttu-id="a5143-433">名前</span><span class="sxs-lookup"><span data-stu-id="a5143-433">Name</span></span>| <span data-ttu-id="a5143-434">型</span><span class="sxs-lookup"><span data-stu-id="a5143-434">Type</span></span>| <span data-ttu-id="a5143-435">属性</span><span class="sxs-lookup"><span data-stu-id="a5143-435">Attributes</span></span>| <span data-ttu-id="a5143-436">説明</span><span class="sxs-lookup"><span data-stu-id="a5143-436">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="a5143-437">Object</span><span class="sxs-lookup"><span data-stu-id="a5143-437">Object</span></span> | <span data-ttu-id="a5143-438">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a5143-438">&lt;optional&gt;</span></span> | <span data-ttu-id="a5143-439">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="a5143-439">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="a5143-440">Boolean</span><span class="sxs-lookup"><span data-stu-id="a5143-440">Boolean</span></span> |  <span data-ttu-id="a5143-441">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="a5143-441">&lt;optional&gt;</span></span> | <span data-ttu-id="a5143-p126">提供されたトークンを Outlook REST API または Exchange Web サービスに使用するかどうかを決定します。既定値は、`false` です。</span><span class="sxs-lookup"><span data-stu-id="a5143-p126">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="a5143-444">Object</span><span class="sxs-lookup"><span data-stu-id="a5143-444">Object</span></span> |  <span data-ttu-id="a5143-445">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="a5143-445">&lt;optional&gt;</span></span> | <span data-ttu-id="a5143-446">非同期メソッドに渡される状態データ。</span><span class="sxs-lookup"><span data-stu-id="a5143-446">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="a5143-447">function</span><span class="sxs-lookup"><span data-stu-id="a5143-447">function</span></span>||<span data-ttu-id="a5143-p127">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="a5143-p127">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a5143-450">要件</span><span class="sxs-lookup"><span data-stu-id="a5143-450">Requirements</span></span>

|<span data-ttu-id="a5143-451">要件</span><span class="sxs-lookup"><span data-stu-id="a5143-451">Requirement</span></span>| <span data-ttu-id="a5143-452">値</span><span class="sxs-lookup"><span data-stu-id="a5143-452">Value</span></span>|
|---|---|
|[<span data-ttu-id="a5143-453">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a5143-453">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a5143-454">1.5</span><span class="sxs-lookup"><span data-stu-id="a5143-454">1.5</span></span> |
|[<span data-ttu-id="a5143-455">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="a5143-455">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a5143-456">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a5143-456">ReadItem</span></span>|
|[<span data-ttu-id="a5143-457">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a5143-457">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a5143-458">新規作成と閲覧</span><span class="sxs-lookup"><span data-stu-id="a5143-458">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="a5143-459">例</span><span class="sxs-lookup"><span data-stu-id="a5143-459">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="a5143-460">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="a5143-460">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="a5143-461">Exchange Server から添付ファイルやアイテムを取得するために使うトークンを含む文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="a5143-461">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="a5143-p128">`getCallbackTokenAsync` メソッドは、ユーザーのメールボックスをホストする Exchange Server から不透明なトークンを取得する非同期の呼び出しを行います。コールバック トークンの有効期間は 5 分です。</span><span class="sxs-lookup"><span data-stu-id="a5143-p128">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="a5143-p129">トークンと予定の識別子またはアイテムの識別子をサードパーティ システムに渡すことができます。サードパーティ システムは、トークンをベアラー承認トークンとして使用し、Exchange Web サービス (EWS) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) または [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) 操作を呼び出して、添付ファイルまたはアイテムを返します。たとえば、[選択したアイテムから添付ファイルを取得する](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="a5143-p129">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="a5143-467">アプリが閲覧モードで `getCallbackTokenAsync` メソッドを呼び出すには、アプリのマニフェスト内に **ReadItem** アクセス許可が指定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="a5143-467">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="a5143-p130">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出してアイテムの識別子を `getCallbackTokenAsync` メソッドに渡す必要があります。アプリには、`saveAsync` メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="a5143-p130">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a5143-470">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="a5143-470">Parameters:</span></span>

|<span data-ttu-id="a5143-471">名前</span><span class="sxs-lookup"><span data-stu-id="a5143-471">Name</span></span>| <span data-ttu-id="a5143-472">型</span><span class="sxs-lookup"><span data-stu-id="a5143-472">Type</span></span>| <span data-ttu-id="a5143-473">属性</span><span class="sxs-lookup"><span data-stu-id="a5143-473">Attributes</span></span>| <span data-ttu-id="a5143-474">説明</span><span class="sxs-lookup"><span data-stu-id="a5143-474">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="a5143-475">function</span><span class="sxs-lookup"><span data-stu-id="a5143-475">function</span></span>||<span data-ttu-id="a5143-p131">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="a5143-p131">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="a5143-478">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="a5143-478">Object</span></span>| <span data-ttu-id="a5143-479">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="a5143-479">&lt;optional&gt;</span></span>|<span data-ttu-id="a5143-480">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="a5143-480">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a5143-481">要件</span><span class="sxs-lookup"><span data-stu-id="a5143-481">Requirements</span></span>

|<span data-ttu-id="a5143-482">要件</span><span class="sxs-lookup"><span data-stu-id="a5143-482">Requirement</span></span>| <span data-ttu-id="a5143-483">値</span><span class="sxs-lookup"><span data-stu-id="a5143-483">Value</span></span>|
|---|---|
|[<span data-ttu-id="a5143-484">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a5143-484">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a5143-485">1.3</span><span class="sxs-lookup"><span data-stu-id="a5143-485">1.3</span></span>|
|[<span data-ttu-id="a5143-486">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="a5143-486">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a5143-487">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a5143-487">ReadItem</span></span>|
|[<span data-ttu-id="a5143-488">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a5143-488">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a5143-489">新規作成と閲覧</span><span class="sxs-lookup"><span data-stu-id="a5143-489">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="a5143-490">例</span><span class="sxs-lookup"><span data-stu-id="a5143-490">Example</span></span>

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="a5143-491">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="a5143-491">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="a5143-492">ユーザーと Office アドインを識別するトークンを取得します。</span><span class="sxs-lookup"><span data-stu-id="a5143-492">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="a5143-493">`getUserIdentityTokenAsync` メソッドは、[アドインとユーザーをサード パーティのシステムで識別して認証](https://docs.microsoft.com/outlook/add-ins/authentication)することのできるトークンを返します。</span><span class="sxs-lookup"><span data-stu-id="a5143-493">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](https://docs.microsoft.com/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="a5143-494">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="a5143-494">Parameters:</span></span>

|<span data-ttu-id="a5143-495">名前</span><span class="sxs-lookup"><span data-stu-id="a5143-495">Name</span></span>| <span data-ttu-id="a5143-496">型</span><span class="sxs-lookup"><span data-stu-id="a5143-496">Type</span></span>| <span data-ttu-id="a5143-497">属性</span><span class="sxs-lookup"><span data-stu-id="a5143-497">Attributes</span></span>| <span data-ttu-id="a5143-498">説明</span><span class="sxs-lookup"><span data-stu-id="a5143-498">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="a5143-499">function</span><span class="sxs-lookup"><span data-stu-id="a5143-499">function</span></span>||<span data-ttu-id="a5143-500">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="a5143-500">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="a5143-501">トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="a5143-501">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="a5143-502">Object</span><span class="sxs-lookup"><span data-stu-id="a5143-502">Object</span></span>| <span data-ttu-id="a5143-503">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="a5143-503">&lt;optional&gt;</span></span>|<span data-ttu-id="a5143-504">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="a5143-504">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a5143-505">要件</span><span class="sxs-lookup"><span data-stu-id="a5143-505">Requirements</span></span>

|<span data-ttu-id="a5143-506">要件</span><span class="sxs-lookup"><span data-stu-id="a5143-506">Requirement</span></span>| <span data-ttu-id="a5143-507">値</span><span class="sxs-lookup"><span data-stu-id="a5143-507">Value</span></span>|
|---|---|
|[<span data-ttu-id="a5143-508">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a5143-508">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a5143-509">1.0</span><span class="sxs-lookup"><span data-stu-id="a5143-509">1.0</span></span>|
|[<span data-ttu-id="a5143-510">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="a5143-510">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a5143-511">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a5143-511">ReadItem</span></span>|
|[<span data-ttu-id="a5143-512">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a5143-512">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a5143-513">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="a5143-513">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a5143-514">例</span><span class="sxs-lookup"><span data-stu-id="a5143-514">Example</span></span>

```js
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="a5143-515">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="a5143-515">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="a5143-516">ユーザーのメールボックスをホストしている Exchange サーバー上の Exchange Web サービス (EWS) のサービスに対して非同期の要求を行います。</span><span class="sxs-lookup"><span data-stu-id="a5143-516">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="a5143-517">このメソッドは、次のシナリオではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="a5143-517">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="a5143-518">Outlook for iOS または Outlook for Android を使用している場合</span><span class="sxs-lookup"><span data-stu-id="a5143-518">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="a5143-519">アドインが Gmail のメールボックスに読み込まれる場合</span><span class="sxs-lookup"><span data-stu-id="a5143-519">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="a5143-520">このような場合は、アドインでは [REST API を使用](https://docs.microsoft.com/outlook/add-ins/use-rest-api)して、代わりにユーザーのメールボックスにアクセスする必要があります。</span><span class="sxs-lookup"><span data-stu-id="a5143-520">In these cases, add-ins should [use REST APIs](https://docs.microsoft.com/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="a5143-521">`makeEwsRequestAsync` メソッドは、アドインの代わりに Exchange に EWS 要求を送信します。</span><span class="sxs-lookup"><span data-stu-id="a5143-521">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="a5143-522">サポートされている EWS 操作の一覧については、「[Outlook アドインからの Web サービスの呼び出し](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="a5143-522">See [Call web services from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="a5143-523">`makeEwsRequestAsync` メソッドでは、フォルダー関連アイテムを要求できません。</span><span class="sxs-lookup"><span data-stu-id="a5143-523">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="a5143-524">XML 要求では UTF-8 エンコードを指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="a5143-524">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="a5143-p133">`makeEwsRequestAsync` メソッドを使用するには、アドインに **ReadWriteMailbox** アクセス許可が必要です。**ReadWriteMailbox** アクセス許可と、`makeEwsRequestAsync` メソッドで呼び出せる EWS 操作の使い方については、「[ユーザーのメールボックスへのメール アドイン アクセスのアクセス許可を指定する](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="a5143-p133">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="a5143-527">サーバー管理者は、クライアント アクセス サーバーの EWS ディレクトリで `OAuthAuthentication` を true に設定して、`makeEwsRequestAsync` メソッドで EWS 要求を行うことができるようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="a5143-527">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="a5143-528">バージョンの相違点</span><span class="sxs-lookup"><span data-stu-id="a5143-528">Version differences</span></span>

<span data-ttu-id="a5143-529">バージョン 15.0.4535.1004 より前のバージョンの Outlook で実行しているメール アプリで `makeEwsRequestAsync` メソッドを使う場合は、エンコード値を `ISO-8859-1` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="a5143-529">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="a5143-p134">Outlook on the web でメール アプリを実行している場合は、エンコード値を設定する必要はありません。mailbox.diagnostics.hostName プロパティを使って、メール アプリを Outlook で実行しているのか、Outlook on the web で実行しているのかを確認できます。mailbox.diagnostics.hostVersion プロパティを使って、どのバージョンの Outlook を使って実行しているかを確認できます。</span><span class="sxs-lookup"><span data-stu-id="a5143-p134">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a5143-533">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="a5143-533">Parameters:</span></span>

|<span data-ttu-id="a5143-534">名前</span><span class="sxs-lookup"><span data-stu-id="a5143-534">Name</span></span>| <span data-ttu-id="a5143-535">型</span><span class="sxs-lookup"><span data-stu-id="a5143-535">Type</span></span>| <span data-ttu-id="a5143-536">属性</span><span class="sxs-lookup"><span data-stu-id="a5143-536">Attributes</span></span>| <span data-ttu-id="a5143-537">説明</span><span class="sxs-lookup"><span data-stu-id="a5143-537">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="a5143-538">String</span><span class="sxs-lookup"><span data-stu-id="a5143-538">String</span></span>||<span data-ttu-id="a5143-539">EWS 要求です。</span><span class="sxs-lookup"><span data-stu-id="a5143-539">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="a5143-540">function</span><span class="sxs-lookup"><span data-stu-id="a5143-540">function</span></span>||<span data-ttu-id="a5143-541">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="a5143-541">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="a5143-542">EWS 呼び出しの XML 結果は、`asyncResult.value` プロパティ内の文字列として提供されています。</span><span class="sxs-lookup"><span data-stu-id="a5143-542">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="a5143-543">結果のサイズが 1 MB を超える場合、代わりにエラー メッセージが返されます。</span><span class="sxs-lookup"><span data-stu-id="a5143-543">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="a5143-544">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="a5143-544">Object</span></span>| <span data-ttu-id="a5143-545">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="a5143-545">&lt;optional&gt;</span></span>|<span data-ttu-id="a5143-546">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="a5143-546">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a5143-547">要件</span><span class="sxs-lookup"><span data-stu-id="a5143-547">Requirements</span></span>

|<span data-ttu-id="a5143-548">要件</span><span class="sxs-lookup"><span data-stu-id="a5143-548">Requirement</span></span>| <span data-ttu-id="a5143-549">値</span><span class="sxs-lookup"><span data-stu-id="a5143-549">Value</span></span>|
|---|---|
|[<span data-ttu-id="a5143-550">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a5143-550">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a5143-551">1.0</span><span class="sxs-lookup"><span data-stu-id="a5143-551">1.0</span></span>|
|[<span data-ttu-id="a5143-552">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="a5143-552">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a5143-553">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="a5143-553">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="a5143-554">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a5143-554">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a5143-555">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="a5143-555">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a5143-556">例</span><span class="sxs-lookup"><span data-stu-id="a5143-556">Example</span></span>

<span data-ttu-id="a5143-557">次の例は、`GetItem` 操作を使ってアイテムの件名を取得するため、`makeEwsRequestAsync` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="a5143-557">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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

####  <a name="removehandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="a5143-558">removeHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="a5143-558">removeHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="a5143-559">サポートされているイベントのイベント ハンドラーを削除します。</span><span class="sxs-lookup"><span data-stu-id="a5143-559">Removes an event handler for a supported event.</span></span>

<span data-ttu-id="a5143-560">現在、サポートされているイベントの種類は `Office.EventType.ItemChanged` だけです。</span><span class="sxs-lookup"><span data-stu-id="a5143-560">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a5143-561">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="a5143-561">Parameters:</span></span>

| <span data-ttu-id="a5143-562">名前</span><span class="sxs-lookup"><span data-stu-id="a5143-562">Name</span></span> | <span data-ttu-id="a5143-563">型</span><span class="sxs-lookup"><span data-stu-id="a5143-563">Type</span></span> | <span data-ttu-id="a5143-564">属性</span><span class="sxs-lookup"><span data-stu-id="a5143-564">Attributes</span></span> | <span data-ttu-id="a5143-565">説明</span><span class="sxs-lookup"><span data-stu-id="a5143-565">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="a5143-566">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="a5143-566">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="a5143-567">ハンドラーを取り消すイベント。</span><span class="sxs-lookup"><span data-stu-id="a5143-567">The event that should revoke the handler.</span></span> |
| `handler` | <span data-ttu-id="a5143-568">職務</span><span class="sxs-lookup"><span data-stu-id="a5143-568">Function</span></span> || <span data-ttu-id="a5143-p136">イベントを処理する関数。関数は、オブジェクト リテラルである単一パラメーターを受け入れる必要があります。パラメーターの `type` プロパティは、`addHandlerAsync` に渡される `eventType` パラメーターと一致します。</span><span class="sxs-lookup"><span data-stu-id="a5143-p136">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="a5143-572">Object</span><span class="sxs-lookup"><span data-stu-id="a5143-572">Object</span></span> | <span data-ttu-id="a5143-573">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a5143-573">&lt;optional&gt;</span></span> | <span data-ttu-id="a5143-574">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="a5143-574">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="a5143-575">Object</span><span class="sxs-lookup"><span data-stu-id="a5143-575">Object</span></span> | <span data-ttu-id="a5143-576">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a5143-576">&lt;optional&gt;</span></span> | <span data-ttu-id="a5143-577">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="a5143-577">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="a5143-578">function</span><span class="sxs-lookup"><span data-stu-id="a5143-578">function</span></span>| <span data-ttu-id="a5143-579">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a5143-579">&lt;optional&gt;</span></span>|<span data-ttu-id="a5143-580">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="a5143-580">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a5143-581">要件</span><span class="sxs-lookup"><span data-stu-id="a5143-581">Requirements</span></span>

|<span data-ttu-id="a5143-582">要件</span><span class="sxs-lookup"><span data-stu-id="a5143-582">Requirement</span></span>| <span data-ttu-id="a5143-583">値</span><span class="sxs-lookup"><span data-stu-id="a5143-583">Value</span></span>|
|---|---|
|[<span data-ttu-id="a5143-584">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a5143-584">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a5143-585">1.5</span><span class="sxs-lookup"><span data-stu-id="a5143-585">1.5</span></span> |
|[<span data-ttu-id="a5143-586">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="a5143-586">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a5143-587">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a5143-587">ReadItem</span></span> |
|[<span data-ttu-id="a5143-588">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a5143-588">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a5143-589">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="a5143-589">Compose or read</span></span>|