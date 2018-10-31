
# <a name="mailbox"></a><span data-ttu-id="7d0dd-101">mailbox</span><span class="sxs-lookup"><span data-stu-id="7d0dd-101">mailbox</span></span>

### <span data-ttu-id="7d0dd-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span><span class="sxs-lookup"><span data-stu-id="7d0dd-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span></span>

<span data-ttu-id="7d0dd-104">Microsoft Outlook と Microsoft Outlook on the web の Outlook アドイン オブジェクト モデルへのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-104">Provides access to the Outlook Add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="7d0dd-105">要件</span><span class="sxs-lookup"><span data-stu-id="7d0dd-105">Requirements</span></span>

|<span data-ttu-id="7d0dd-106">必要条件</span><span class="sxs-lookup"><span data-stu-id="7d0dd-106">Requirement</span></span>| <span data-ttu-id="7d0dd-107">値</span><span class="sxs-lookup"><span data-stu-id="7d0dd-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="7d0dd-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="7d0dd-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7d0dd-109">1.0以降</span><span class="sxs-lookup"><span data-stu-id="7d0dd-109">1.0</span></span>|
|[<span data-ttu-id="7d0dd-110">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="7d0dd-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7d0dd-111">制限あり</span><span class="sxs-lookup"><span data-stu-id="7d0dd-111">Restricted</span></span>|
|[<span data-ttu-id="7d0dd-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="7d0dd-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7d0dd-113">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="7d0dd-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="7d0dd-114">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="7d0dd-114">Members and methods</span></span>

| <span data-ttu-id="7d0dd-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="7d0dd-115">Member</span></span> | <span data-ttu-id="7d0dd-116">種類</span><span class="sxs-lookup"><span data-stu-id="7d0dd-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="7d0dd-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="7d0dd-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="7d0dd-118">メンバー</span><span class="sxs-lookup"><span data-stu-id="7d0dd-118">Member</span></span> |
| [<span data-ttu-id="7d0dd-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="7d0dd-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="7d0dd-120">メンバー</span><span class="sxs-lookup"><span data-stu-id="7d0dd-120">Member</span></span> |
| [<span data-ttu-id="7d0dd-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="7d0dd-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="7d0dd-122">メソッド</span><span class="sxs-lookup"><span data-stu-id="7d0dd-122">Method</span></span> |
| [<span data-ttu-id="7d0dd-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="7d0dd-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="7d0dd-124">メソッド</span><span class="sxs-lookup"><span data-stu-id="7d0dd-124">Method</span></span> |
| [<span data-ttu-id="7d0dd-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="7d0dd-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) | <span data-ttu-id="7d0dd-126">メソッド</span><span class="sxs-lookup"><span data-stu-id="7d0dd-126">Method</span></span> |
| [<span data-ttu-id="7d0dd-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="7d0dd-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="7d0dd-128">メソッド</span><span class="sxs-lookup"><span data-stu-id="7d0dd-128">Method</span></span> |
| [<span data-ttu-id="7d0dd-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="7d0dd-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="7d0dd-130">メソッド</span><span class="sxs-lookup"><span data-stu-id="7d0dd-130">Method</span></span> |
| [<span data-ttu-id="7d0dd-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="7d0dd-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="7d0dd-132">メソッド</span><span class="sxs-lookup"><span data-stu-id="7d0dd-132">Method</span></span> |
| [<span data-ttu-id="7d0dd-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="7d0dd-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="7d0dd-134">メソッド</span><span class="sxs-lookup"><span data-stu-id="7d0dd-134">Method</span></span> |
| [<span data-ttu-id="7d0dd-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="7d0dd-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="7d0dd-136">メソッド</span><span class="sxs-lookup"><span data-stu-id="7d0dd-136">Method</span></span> |
| [<span data-ttu-id="7d0dd-137">displayNewMessageForm</span><span class="sxs-lookup"><span data-stu-id="7d0dd-137">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="7d0dd-138">メソッド</span><span class="sxs-lookup"><span data-stu-id="7d0dd-138">Method</span></span> |
| [<span data-ttu-id="7d0dd-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="7d0dd-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="7d0dd-140">メソッド</span><span class="sxs-lookup"><span data-stu-id="7d0dd-140">Method</span></span> |
| [<span data-ttu-id="7d0dd-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="7d0dd-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="7d0dd-142">メソッド</span><span class="sxs-lookup"><span data-stu-id="7d0dd-142">Method</span></span> |
| [<span data-ttu-id="7d0dd-143">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="7d0dd-143">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="7d0dd-144">メソッド</span><span class="sxs-lookup"><span data-stu-id="7d0dd-144">Method</span></span> |
| [<span data-ttu-id="7d0dd-145">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="7d0dd-145">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="7d0dd-146">メソッド</span><span class="sxs-lookup"><span data-stu-id="7d0dd-146">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="7d0dd-147">名前空間</span><span class="sxs-lookup"><span data-stu-id="7d0dd-147">Namespaces</span></span>

<span data-ttu-id="7d0dd-148">[diagnostics](Office.context.mailbox.diagnostics.md): Outlook アドインに診断情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-148">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="7d0dd-149">[item](Office.context.mailbox.item.md):Outlook アドインのメッセージや予定にアクセスするためのメソッドとプロパティを提供します。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-149">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="7d0dd-150">[userProfile](Office.context.mailbox.userProfile.md):Outlook アドインのユーザーに関する情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-150">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="7d0dd-151">メンバー</span><span class="sxs-lookup"><span data-stu-id="7d0dd-151">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="7d0dd-152">ewsUrl: 文字列</span><span class="sxs-lookup"><span data-stu-id="7d0dd-152">ewsUrl :String</span></span>

<span data-ttu-id="7d0dd-p102">このメール アカウントの Exchange Web サービス (EWS) エンドポイントの URL を取得します。閲覧モードのみです。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-p102">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="7d0dd-155">このメンバーは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-155">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="7d0dd-p103">`ewsUrl` 値は、リモート サービスで、ユーザーのメールボックスに EWS 呼び出しを行うために使用することができます。たとえば、[選択した項目から添付ファイルを取得する](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-p103">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="7d0dd-158">閲覧モードで `ewsUrl` メンバーを呼び出すには、アプリでマニフェスト内に **ReadItem** アクセス許可を指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-158">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="7d0dd-p104">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback)メソッドを呼び出してから、`ewsUrl`メンバーを使用する必要があります。アプリには、`saveAsync`メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="7d0dd-161">種類:</span><span class="sxs-lookup"><span data-stu-id="7d0dd-161">Type:</span></span>

*   <span data-ttu-id="7d0dd-162">文字列</span><span class="sxs-lookup"><span data-stu-id="7d0dd-162">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="7d0dd-163">要件</span><span class="sxs-lookup"><span data-stu-id="7d0dd-163">Requirements</span></span>

|<span data-ttu-id="7d0dd-164">必要条件</span><span class="sxs-lookup"><span data-stu-id="7d0dd-164">Requirement</span></span>| <span data-ttu-id="7d0dd-165">値</span><span class="sxs-lookup"><span data-stu-id="7d0dd-165">Value</span></span>|
|---|---|
|[<span data-ttu-id="7d0dd-166">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="7d0dd-166">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7d0dd-167">1.0以降</span><span class="sxs-lookup"><span data-stu-id="7d0dd-167">1.0</span></span>|
|[<span data-ttu-id="7d0dd-168">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="7d0dd-168">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7d0dd-169">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7d0dd-169">ReadItem</span></span>|
|[<span data-ttu-id="7d0dd-170">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="7d0dd-170">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7d0dd-171">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="7d0dd-171">Compose or read</span></span>|

#### <a name="resturl-string"></a><span data-ttu-id="7d0dd-172">restUrl: 文字列</span><span class="sxs-lookup"><span data-stu-id="7d0dd-172">restUrl :String</span></span>

<span data-ttu-id="7d0dd-173">この電子メール アカウントの REST エンドポイントの URL を取得します。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-173">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="7d0dd-174">`restUrl` 値は、ユーザーのメールボックスに [REST API](https://docs.microsoft.com/outlook/rest/) 呼び出しを行うために使用することができます。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-174">The `restUrl` value can be used to make [REST API](https://docs.microsoft.com/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="7d0dd-175">閲覧モードで `restUrl` メンバーを呼び出すには、アプリでマニフェスト内に **ReadItem** アクセス許可を指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-175">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="7d0dd-p105">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback)メソッドを呼び出してから、`restUrl`メンバーを使用する必要があります。アプリには、`saveAsync`メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-p105">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="7d0dd-178">種類:</span><span class="sxs-lookup"><span data-stu-id="7d0dd-178">Type:</span></span>

*   <span data-ttu-id="7d0dd-179">文字列</span><span class="sxs-lookup"><span data-stu-id="7d0dd-179">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="7d0dd-180">要件</span><span class="sxs-lookup"><span data-stu-id="7d0dd-180">Requirements</span></span>

|<span data-ttu-id="7d0dd-181">必要条件</span><span class="sxs-lookup"><span data-stu-id="7d0dd-181">Requirement</span></span>| <span data-ttu-id="7d0dd-182">値</span><span class="sxs-lookup"><span data-stu-id="7d0dd-182">Value</span></span>|
|---|---|
|[<span data-ttu-id="7d0dd-183">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="7d0dd-183">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7d0dd-184">1.5</span><span class="sxs-lookup"><span data-stu-id="7d0dd-184">1.5</span></span> |
|[<span data-ttu-id="7d0dd-185">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="7d0dd-185">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7d0dd-186">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7d0dd-186">ReadItem</span></span>|
|[<span data-ttu-id="7d0dd-187">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="7d0dd-187">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7d0dd-188">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="7d0dd-188">Compose or read</span></span>|

### <a name="methods"></a><span data-ttu-id="7d0dd-189">メソッド</span><span class="sxs-lookup"><span data-stu-id="7d0dd-189">Methods</span></span>

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="7d0dd-190">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="7d0dd-190">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="7d0dd-191">サポートされているイベントのイベント ハンドラーを追加します。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-191">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="7d0dd-p106">現在サポートされている唯一のイベントの種類は `Office.EventType.ItemChanged` です。これはユーザーが新しいアイテムを選択したときに呼び出されます。このイベントは、ピン留め可能な作業ウィンドウを実装するアドインで使用され、現在選択されているアイテムに基づいて作業ウィンドウ UI をアドインで更新できるようにします。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-p106">Currently the only supported event type is `Office.EventType.ItemChanged`, which is invoked when the user selects a new item. This event is used by add-ins that implement a pinnable taskpane, and allows the add-in to refresh the taskpane UI based on the currently selected item.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7d0dd-194">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="7d0dd-194">Parameters:</span></span>

| <span data-ttu-id="7d0dd-195">名前</span><span class="sxs-lookup"><span data-stu-id="7d0dd-195">Name</span></span> | <span data-ttu-id="7d0dd-196">種類</span><span class="sxs-lookup"><span data-stu-id="7d0dd-196">Type</span></span> | <span data-ttu-id="7d0dd-197">属性</span><span class="sxs-lookup"><span data-stu-id="7d0dd-197">Attributes</span></span> | <span data-ttu-id="7d0dd-198">説明</span><span class="sxs-lookup"><span data-stu-id="7d0dd-198">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="7d0dd-199">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="7d0dd-199">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="7d0dd-200">ハンドラを呼び出す必要のあるイベント。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-200">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="7d0dd-201">関数</span><span class="sxs-lookup"><span data-stu-id="7d0dd-201">Function</span></span> || <span data-ttu-id="7d0dd-p107">イベントを処理する関数。この関数は、オブジェクト リテラルである単一パラメータを受け入れる必要があります。パラメータの `type` プロパティは、`addHandlerAsync` に渡される `eventType` パラメータと一致します。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-p107">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="7d0dd-205">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="7d0dd-205">Object</span></span> | <span data-ttu-id="7d0dd-206">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="7d0dd-206">&lt;optional&gt;</span></span> | <span data-ttu-id="7d0dd-207">次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-207">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="7d0dd-208">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="7d0dd-208">Object</span></span> | <span data-ttu-id="7d0dd-209">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="7d0dd-209">&lt;optional&gt;</span></span> | <span data-ttu-id="7d0dd-210">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-210">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="7d0dd-211">関数</span><span class="sxs-lookup"><span data-stu-id="7d0dd-211">function</span></span>| <span data-ttu-id="7d0dd-212">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="7d0dd-212">&lt;optional&gt;</span></span>|<span data-ttu-id="7d0dd-213">メソッドが完了すると、`callback` パラメータに渡された関数が、単一パラメータ `asyncResult` で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-213">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7d0dd-214">要件</span><span class="sxs-lookup"><span data-stu-id="7d0dd-214">Requirements</span></span>

|<span data-ttu-id="7d0dd-215">必要条件</span><span class="sxs-lookup"><span data-stu-id="7d0dd-215">Requirement</span></span>| <span data-ttu-id="7d0dd-216">値</span><span class="sxs-lookup"><span data-stu-id="7d0dd-216">Value</span></span>|
|---|---|
|[<span data-ttu-id="7d0dd-217">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="7d0dd-217">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7d0dd-218">1.5</span><span class="sxs-lookup"><span data-stu-id="7d0dd-218">1.5</span></span> |
|[<span data-ttu-id="7d0dd-219">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="7d0dd-219">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7d0dd-220">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7d0dd-220">ReadItem</span></span> |
|[<span data-ttu-id="7d0dd-221">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="7d0dd-221">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7d0dd-222">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="7d0dd-222">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="7d0dd-223">例</span><span class="sxs-lookup"><span data-stu-id="7d0dd-223">Example</span></span>

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

####  <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="7d0dd-224">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="7d0dd-224">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="7d0dd-225">REST 用に書式設定された項目 ID を EWS 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-225">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="7d0dd-226">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-226">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="7d0dd-p108">REST API ([Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) や [Microsoft Graph](http://graph.microsoft.io/) など) 経由で取得された項目 ID は、Exchange Web サービス (EWS) で使用される形式とは異なる形式を使用します。`convertToEwsId` メソッドは、REST 形式の ID を EWS 用の適切な形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-p108">Item IDs retrieved via a REST API (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](http://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7d0dd-229">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="7d0dd-229">Parameters:</span></span>

|<span data-ttu-id="7d0dd-230">名前</span><span class="sxs-lookup"><span data-stu-id="7d0dd-230">Name</span></span>| <span data-ttu-id="7d0dd-231">種類</span><span class="sxs-lookup"><span data-stu-id="7d0dd-231">Type</span></span>| <span data-ttu-id="7d0dd-232">説明</span><span class="sxs-lookup"><span data-stu-id="7d0dd-232">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="7d0dd-233">文字列</span><span class="sxs-lookup"><span data-stu-id="7d0dd-233">String</span></span>|<span data-ttu-id="7d0dd-234">Outlook REST API 用に書式設定された項目 ID</span><span class="sxs-lookup"><span data-stu-id="7d0dd-234">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="7d0dd-235">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="7d0dd-235">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.restversion)|<span data-ttu-id="7d0dd-236">項目 ID の取得に使用された Outlook REST API のバージョンを示す値。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-236">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7d0dd-237">要件</span><span class="sxs-lookup"><span data-stu-id="7d0dd-237">Requirements</span></span>

|<span data-ttu-id="7d0dd-238">必要条件</span><span class="sxs-lookup"><span data-stu-id="7d0dd-238">Requirement</span></span>| <span data-ttu-id="7d0dd-239">値</span><span class="sxs-lookup"><span data-stu-id="7d0dd-239">Value</span></span>|
|---|---|
|[<span data-ttu-id="7d0dd-240">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="7d0dd-240">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7d0dd-241">1.3</span><span class="sxs-lookup"><span data-stu-id="7d0dd-241">1.3</span></span>|
|[<span data-ttu-id="7d0dd-242">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="7d0dd-242">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7d0dd-243">制限あり</span><span class="sxs-lookup"><span data-stu-id="7d0dd-243">Restricted</span></span>|
|[<span data-ttu-id="7d0dd-244">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="7d0dd-244">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7d0dd-245">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="7d0dd-245">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="7d0dd-246">戻り値 :</span><span class="sxs-lookup"><span data-stu-id="7d0dd-246">Returns:</span></span>

<span data-ttu-id="7d0dd-247">種類: 文字列</span><span class="sxs-lookup"><span data-stu-id="7d0dd-247">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="7d0dd-248">例</span><span class="sxs-lookup"><span data-stu-id="7d0dd-248">Example</span></span>

```
// Get an item's ID from a REST API
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the
// Outlook Mail API
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime"></a><span data-ttu-id="7d0dd-249">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="7d0dd-249">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)}</span></span>

<span data-ttu-id="7d0dd-250">クライアントの現地時間の時間情報が含まれている辞書を取得します。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-250">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="7d0dd-p109">Outlook 用メール アプリや Outlook Web App で使う日付と時刻は、異なるタイム ゾーンを使うことができます。Outlook では、クライアント コンピューターのタイム ゾーンを使います。Outlook Web App では、Exchange 管理センター (EAC) で設定されたタイム ゾーンを使います。ユーザー インターフェイスに表示される値が、常にユーザーが期待するタイム ゾーンと一致するように日付と時刻の値を処理する必要があります。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-p109">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="7d0dd-p110">Outlook でメール アプリが実行されている場合、`convertToLocalClientTime` メソッドは、クライアント コンピューターのタイム ゾーンに設定された値で辞書オブジェクトを返します。Outlook Web Apps でメール アプリが実行されている場合、`convertToLocalClientTime` メソッドは、EAC で指定されたタイム ゾーンに設定された値で辞書オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-p110">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7d0dd-256">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="7d0dd-256">Parameters:</span></span>

|<span data-ttu-id="7d0dd-257">名前</span><span class="sxs-lookup"><span data-stu-id="7d0dd-257">Name</span></span>| <span data-ttu-id="7d0dd-258">種類</span><span class="sxs-lookup"><span data-stu-id="7d0dd-258">Type</span></span>| <span data-ttu-id="7d0dd-259">説明</span><span class="sxs-lookup"><span data-stu-id="7d0dd-259">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="7d0dd-260">日付</span><span class="sxs-lookup"><span data-stu-id="7d0dd-260">Date</span></span>|<span data-ttu-id="7d0dd-261">Date オブジェクト</span><span class="sxs-lookup"><span data-stu-id="7d0dd-261">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7d0dd-262">要件</span><span class="sxs-lookup"><span data-stu-id="7d0dd-262">Requirements</span></span>

|<span data-ttu-id="7d0dd-263">必要条件</span><span class="sxs-lookup"><span data-stu-id="7d0dd-263">Requirement</span></span>| <span data-ttu-id="7d0dd-264">値</span><span class="sxs-lookup"><span data-stu-id="7d0dd-264">Value</span></span>|
|---|---|
|[<span data-ttu-id="7d0dd-265">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="7d0dd-265">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7d0dd-266">1.0以降</span><span class="sxs-lookup"><span data-stu-id="7d0dd-266">1.0</span></span>|
|[<span data-ttu-id="7d0dd-267">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="7d0dd-267">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7d0dd-268">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7d0dd-268">ReadItem</span></span>|
|[<span data-ttu-id="7d0dd-269">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="7d0dd-269">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7d0dd-270">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="7d0dd-270">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="7d0dd-271">戻り値 :</span><span class="sxs-lookup"><span data-stu-id="7d0dd-271">Returns:</span></span>

<span data-ttu-id="7d0dd-272">種類:[LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="7d0dd-272">Type: [LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)</span></span>

####  <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="7d0dd-273">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="7d0dd-273">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="7d0dd-274">EWS 用に書式設定された項目 ID を REST 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-274">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="7d0dd-275">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-275">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="7d0dd-p111">EWS 経由または `itemId` プロパティ経由で取得される項目 ID では、REST API ([Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) または [Microsoft Graph](http://graph.microsoft.io/) など) で使用される形式とは異なる形式を使用します。`convertToRestId` メソッドは、EWS 形式の ID を REST 用の適切な形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-p111">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](http://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7d0dd-278">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="7d0dd-278">Parameters:</span></span>

|<span data-ttu-id="7d0dd-279">名前</span><span class="sxs-lookup"><span data-stu-id="7d0dd-279">Name</span></span>| <span data-ttu-id="7d0dd-280">種類</span><span class="sxs-lookup"><span data-stu-id="7d0dd-280">Type</span></span>| <span data-ttu-id="7d0dd-281">説明</span><span class="sxs-lookup"><span data-stu-id="7d0dd-281">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="7d0dd-282">文字列</span><span class="sxs-lookup"><span data-stu-id="7d0dd-282">String</span></span>|<span data-ttu-id="7d0dd-283">Exchange Web サービス (EWS) 用に書式設定された項目 ID</span><span class="sxs-lookup"><span data-stu-id="7d0dd-283">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="7d0dd-284">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="7d0dd-284">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.restversion)|<span data-ttu-id="7d0dd-285">変換後の ID とともに使用される Outlook REST API のバージョンを示す値。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-285">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7d0dd-286">要件</span><span class="sxs-lookup"><span data-stu-id="7d0dd-286">Requirements</span></span>

|<span data-ttu-id="7d0dd-287">必要条件</span><span class="sxs-lookup"><span data-stu-id="7d0dd-287">Requirement</span></span>| <span data-ttu-id="7d0dd-288">値</span><span class="sxs-lookup"><span data-stu-id="7d0dd-288">Value</span></span>|
|---|---|
|[<span data-ttu-id="7d0dd-289">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="7d0dd-289">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7d0dd-290">1.3</span><span class="sxs-lookup"><span data-stu-id="7d0dd-290">1.3</span></span>|
|[<span data-ttu-id="7d0dd-291">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="7d0dd-291">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7d0dd-292">制限あり</span><span class="sxs-lookup"><span data-stu-id="7d0dd-292">Restricted</span></span>|
|[<span data-ttu-id="7d0dd-293">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="7d0dd-293">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7d0dd-294">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="7d0dd-294">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="7d0dd-295">戻り値 :</span><span class="sxs-lookup"><span data-stu-id="7d0dd-295">Returns:</span></span>

<span data-ttu-id="7d0dd-296">種類: 文字列</span><span class="sxs-lookup"><span data-stu-id="7d0dd-296">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="7d0dd-297">例</span><span class="sxs-lookup"><span data-stu-id="7d0dd-297">Example</span></span>

```
// Get the currently selected item's ID
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the
// Outlook Mail API
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="7d0dd-298">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="7d0dd-298">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="7d0dd-299">時間情報が含まれている辞書から Date オブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-299">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="7d0dd-300">`convertToUtcClientTime` メソッドは、ローカルの日付と時刻が含まれる辞書を、ローカルの日付と時刻の正しい値をもつ Date オブジェクトに変換します。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-300">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7d0dd-301">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="7d0dd-301">Parameters:</span></span>

|<span data-ttu-id="7d0dd-302">名前</span><span class="sxs-lookup"><span data-stu-id="7d0dd-302">Name</span></span>| <span data-ttu-id="7d0dd-303">種類</span><span class="sxs-lookup"><span data-stu-id="7d0dd-303">Type</span></span>| <span data-ttu-id="7d0dd-304">説明</span><span class="sxs-lookup"><span data-stu-id="7d0dd-304">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="7d0dd-305">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="7d0dd-305">LocalClientTime</span></span>](/javascript/api/outlook_1_6/office.LocalClientTime)|<span data-ttu-id="7d0dd-306">変換するローカル時刻の値。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-306">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7d0dd-307">要件</span><span class="sxs-lookup"><span data-stu-id="7d0dd-307">Requirements</span></span>

|<span data-ttu-id="7d0dd-308">必要条件</span><span class="sxs-lookup"><span data-stu-id="7d0dd-308">Requirement</span></span>| <span data-ttu-id="7d0dd-309">値</span><span class="sxs-lookup"><span data-stu-id="7d0dd-309">Value</span></span>|
|---|---|
|[<span data-ttu-id="7d0dd-310">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="7d0dd-310">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7d0dd-311">1.0以降</span><span class="sxs-lookup"><span data-stu-id="7d0dd-311">1.0</span></span>|
|[<span data-ttu-id="7d0dd-312">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="7d0dd-312">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7d0dd-313">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7d0dd-313">ReadItem</span></span>|
|[<span data-ttu-id="7d0dd-314">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="7d0dd-314">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7d0dd-315">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="7d0dd-315">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="7d0dd-316">戻り値 :</span><span class="sxs-lookup"><span data-stu-id="7d0dd-316">Returns:</span></span>

<span data-ttu-id="7d0dd-317">時間が UTC で表現された Date オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-317">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="7d0dd-318">

<dt>種類</dt>

</span><span class="sxs-lookup"><span data-stu-id="7d0dd-318">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="7d0dd-319">日付</span><span class="sxs-lookup"><span data-stu-id="7d0dd-319">Date</span></span></dd>

</dl>

####  <a name="displayappointmentformitemid"></a><span data-ttu-id="7d0dd-320">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="7d0dd-320">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="7d0dd-321">既存の予定表の予定を表示します。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-321">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="7d0dd-322">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-322">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="7d0dd-323">`displayAppointmentForm` メソッドは、デスクトップ上の新しいウィンドウまたはモバイル デバイス上のダイアログ ボックスで、既存の予定表の予定を開きます。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-323">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="7d0dd-p112">Outlook for Mac では、このメソッドを使用して、定期的に繰り返される予定の一部ではない単発の予定、または定期的に繰り替えされる予定の元の予定を表示できます。ただし、一連の予定のインスタンスは表示できません。これは、Outlook for Mac では、定期的に繰り返されるインスタンスのプロパティ  (項目 ID を含む) にアクセスできないためです。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-p112">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="7d0dd-326">Outlook Web App では、このメソッドで、指定されたフォームの本文が 32 KB 以下の文字数の場合に、そのフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-326">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="7d0dd-327">指定した項目識別子が既存の予定を識別しない場合には、クライアント コンピュータまたはデバイスで空白のウィンドウが開き、エラー メッセージは返されません。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-327">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7d0dd-328">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="7d0dd-328">Parameters:</span></span>

|<span data-ttu-id="7d0dd-329">名前</span><span class="sxs-lookup"><span data-stu-id="7d0dd-329">Name</span></span>| <span data-ttu-id="7d0dd-330">種類</span><span class="sxs-lookup"><span data-stu-id="7d0dd-330">Type</span></span>| <span data-ttu-id="7d0dd-331">説明</span><span class="sxs-lookup"><span data-stu-id="7d0dd-331">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="7d0dd-332">文字列</span><span class="sxs-lookup"><span data-stu-id="7d0dd-332">String</span></span>|<span data-ttu-id="7d0dd-333">既存の予定表の予定の Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-333">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7d0dd-334">要件</span><span class="sxs-lookup"><span data-stu-id="7d0dd-334">Requirements</span></span>

|<span data-ttu-id="7d0dd-335">必要条件</span><span class="sxs-lookup"><span data-stu-id="7d0dd-335">Requirement</span></span>| <span data-ttu-id="7d0dd-336">値</span><span class="sxs-lookup"><span data-stu-id="7d0dd-336">Value</span></span>|
|---|---|
|[<span data-ttu-id="7d0dd-337">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="7d0dd-337">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7d0dd-338">1.0以降</span><span class="sxs-lookup"><span data-stu-id="7d0dd-338">1.0</span></span>|
|[<span data-ttu-id="7d0dd-339">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="7d0dd-339">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7d0dd-340">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7d0dd-340">ReadItem</span></span>|
|[<span data-ttu-id="7d0dd-341">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="7d0dd-341">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7d0dd-342">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="7d0dd-342">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="7d0dd-343">例</span><span class="sxs-lookup"><span data-stu-id="7d0dd-343">Example</span></span>

```
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  <a name="displaymessageformitemid"></a><span data-ttu-id="7d0dd-344">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="7d0dd-344">displayMessageForm(itemId)</span></span>

<span data-ttu-id="7d0dd-345">既存のメッセージを表示します。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-345">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="7d0dd-346">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-346">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="7d0dd-347">`displayMessageForm` メソッドは、デスクトップ上の新しいウィンドウまたはモバイル デバイス上のダイアログ ボックスで既存のメッセージを開きます。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-347">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="7d0dd-348">Outlook Web App では、このメソッドは、指定されたフォームの本文が 32 KB 以下の文字数の場合に、そのフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-348">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="7d0dd-349">指定した項目識別子が既存のメッセージを識別しない場合には、クラアント コンピュータでメッセージは表示されず、エラー メッセージは返されません。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-349">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="7d0dd-p113">予定を表す `itemId` が含まれる `displayMessageForm` を使用しないでください。`displayAppointmentForm` メソッドを使用して既存の予定を表示し、`displayNewAppointmentForm` を使用して新しい予定を作成するフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-p113">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7d0dd-352">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="7d0dd-352">Parameters:</span></span>

|<span data-ttu-id="7d0dd-353">名前</span><span class="sxs-lookup"><span data-stu-id="7d0dd-353">Name</span></span>| <span data-ttu-id="7d0dd-354">種類</span><span class="sxs-lookup"><span data-stu-id="7d0dd-354">Type</span></span>| <span data-ttu-id="7d0dd-355">説明</span><span class="sxs-lookup"><span data-stu-id="7d0dd-355">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="7d0dd-356">文字列</span><span class="sxs-lookup"><span data-stu-id="7d0dd-356">String</span></span>|<span data-ttu-id="7d0dd-357">既存のメッセージの Exchange Web サービス(EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-357">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7d0dd-358">要件</span><span class="sxs-lookup"><span data-stu-id="7d0dd-358">Requirements</span></span>

|<span data-ttu-id="7d0dd-359">必要条件</span><span class="sxs-lookup"><span data-stu-id="7d0dd-359">Requirement</span></span>| <span data-ttu-id="7d0dd-360">値</span><span class="sxs-lookup"><span data-stu-id="7d0dd-360">Value</span></span>|
|---|---|
|[<span data-ttu-id="7d0dd-361">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="7d0dd-361">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7d0dd-362">1.0以降</span><span class="sxs-lookup"><span data-stu-id="7d0dd-362">1.0</span></span>|
|[<span data-ttu-id="7d0dd-363">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="7d0dd-363">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7d0dd-364">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7d0dd-364">ReadItem</span></span>|
|[<span data-ttu-id="7d0dd-365">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="7d0dd-365">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7d0dd-366">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="7d0dd-366">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="7d0dd-367">例</span><span class="sxs-lookup"><span data-stu-id="7d0dd-367">Example</span></span>

```
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="7d0dd-368">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="7d0dd-368">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="7d0dd-369">新しい予定表の予定を作成するためのフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-369">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="7d0dd-370">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-370">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="7d0dd-p114">`displayNewAppointmentForm` メソッドを使用すると、ユーザーが新しい予定または会議を作成できるフォームが開きます。パラメータを指定すると、予定のフォーム フィールドにパラメータの内容が自動的に入力されます。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-p114">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="7d0dd-p115">Outlook Web App および OWA for Devices では、このメソッドは出席者フィールドが含まれるフォームを常に表示します。入力因数として出席者を指定しない場合には、このメソッドは [**保存**] ボタンのあるフォームを表示します。出席者を指定した場合には、フォームにはその出席者と [**送信**] ボタンが含まれます。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-p115">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="7d0dd-p116">Outlook リッチ クライアントおよび Outlook RT では、`requiredAttendees`、`optionalAttendees` または `resources` パラメータに出席者またはリソースを指定した場合、このメソッドは [**送信**] ボタンがある会議フォームを表示します。受信者を指定しない場合には、このメソッドは [**保存して閉じる**] ボタンがある予定フォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-p116">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="7d0dd-378">パラメータのいずれかが指定されたサイズ制限を超えた場合、または不明なパラメータ名が指定された場合には、例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-378">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7d0dd-379">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="7d0dd-379">Parameters:</span></span>

> [!NOTE]
> <span data-ttu-id="7d0dd-380">すべてのパラメータは省略可能です。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-380">Note: All parameters are optional.</span></span>

|<span data-ttu-id="7d0dd-381">名前</span><span class="sxs-lookup"><span data-stu-id="7d0dd-381">Name</span></span>| <span data-ttu-id="7d0dd-382">種類</span><span class="sxs-lookup"><span data-stu-id="7d0dd-382">Type</span></span>| <span data-ttu-id="7d0dd-383">説明</span><span class="sxs-lookup"><span data-stu-id="7d0dd-383">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="7d0dd-384">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="7d0dd-384">Object</span></span> | <span data-ttu-id="7d0dd-385">新しい予定を記述するパラメータの辞書。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-385">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="7d0dd-386">配列。&lt;文字列&gt; | 配列。&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="7d0dd-386">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="7d0dd-p117">予定への各必須出席者の電子メール アドレスを含む文字列の配列、または出席者の `EmailAddressDetails` オブジェクトを含む配列。この配列は、最大 100 エントリまでに制限されています。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-p117">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="7d0dd-389">配列。&lt;文字列&gt; | 配列。&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="7d0dd-389">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="7d0dd-p118">予定への各任意出席者の電子メール アドレスを含む文字列の配列、または出席者の `EmailAddressDetails` オブジェクトを含む配列。この配列は、最大 100 エントリまでに制限されています。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-p118">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="7d0dd-392">日付</span><span class="sxs-lookup"><span data-stu-id="7d0dd-392">Date</span></span> | <span data-ttu-id="7d0dd-393">予定の開始日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-393">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="7d0dd-394">日付</span><span class="sxs-lookup"><span data-stu-id="7d0dd-394">Date</span></span> | <span data-ttu-id="7d0dd-395">予定の終了日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-395">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="7d0dd-396">文字列</span><span class="sxs-lookup"><span data-stu-id="7d0dd-396">String</span></span> | <span data-ttu-id="7d0dd-p119">予定の場所を含む文字列。文字列は最大 255 文字までに制限されています。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-p119">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="7d0dd-399">配列。&lt; 文字列&gt;</span><span class="sxs-lookup"><span data-stu-id="7d0dd-399">Array.&lt;String&gt;</span></span> | <span data-ttu-id="7d0dd-p120">予定に必要なリソースを含む文字列の配列。配列は最大 100 エントリまでに制限されています。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-p120">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="7d0dd-402">文字列</span><span class="sxs-lookup"><span data-stu-id="7d0dd-402">String</span></span> | <span data-ttu-id="7d0dd-p121">予定の件名を含む文字列。文字列は最大 255 文字までに制限されています。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-p121">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="7d0dd-405">文字列</span><span class="sxs-lookup"><span data-stu-id="7d0dd-405">String</span></span> | <span data-ttu-id="7d0dd-p122">予定の本文。本文の内容は、最大サイズが 32 KB までに制限されています。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-p122">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="7d0dd-408">要件</span><span class="sxs-lookup"><span data-stu-id="7d0dd-408">Requirements</span></span>

|<span data-ttu-id="7d0dd-409">必要条件</span><span class="sxs-lookup"><span data-stu-id="7d0dd-409">Requirement</span></span>| <span data-ttu-id="7d0dd-410">値</span><span class="sxs-lookup"><span data-stu-id="7d0dd-410">Value</span></span>|
|---|---|
|[<span data-ttu-id="7d0dd-411">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="7d0dd-411">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7d0dd-412">1.0以降</span><span class="sxs-lookup"><span data-stu-id="7d0dd-412">1.0</span></span>|
|[<span data-ttu-id="7d0dd-413">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="7d0dd-413">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7d0dd-414">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7d0dd-414">ReadItem</span></span>|
|[<span data-ttu-id="7d0dd-415">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="7d0dd-415">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7d0dd-416">読み取り</span><span class="sxs-lookup"><span data-stu-id="7d0dd-416">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7d0dd-417">例</span><span class="sxs-lookup"><span data-stu-id="7d0dd-417">Example</span></span>

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

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="7d0dd-418">displayNewMessageForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="7d0dd-418">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="7d0dd-419">新しいメッセージを作成するためのフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-419">Displays a form for creating a new message.</span></span>

<span data-ttu-id="7d0dd-420">`displayNewMessageForm` メソッドは、ユーザーが新しいメッセージを作成できるフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-420">The `displayNewMessageForm` method opens a form that enables the user to create a new message.</span></span> <span data-ttu-id="7d0dd-421">パラメータを指定すると、メッセージ フォーム フィールドにはパラメータのコンテンツが自動的に入力されます。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-421">If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="7d0dd-422">パラメータのいずれかが指定されたサイズ制限を超えた場合、または不明なパラメータ名が指定された場合には、例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-422">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7d0dd-423">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="7d0dd-423">Parameters:</span></span>

> [!NOTE]
> <span data-ttu-id="7d0dd-424">すべてのパラメータは省略可能です。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-424">Note: All parameters are optional.</span></span>

|<span data-ttu-id="7d0dd-425">名前</span><span class="sxs-lookup"><span data-stu-id="7d0dd-425">Name</span></span>| <span data-ttu-id="7d0dd-426">種類</span><span class="sxs-lookup"><span data-stu-id="7d0dd-426">Type</span></span>| <span data-ttu-id="7d0dd-427">説明</span><span class="sxs-lookup"><span data-stu-id="7d0dd-427">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="7d0dd-428">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="7d0dd-428">Object</span></span> | <span data-ttu-id="7d0dd-429">新しいメッセージを記述するパラメータの辞書。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-429">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="7d0dd-430">配列。&lt;文字列&gt; | 配列。&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="7d0dd-430">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="7d0dd-431">電子メール アドレスを含む文字列の配列、または To 行の各受信者の `EmailAddressDetails` オブジェクトを含む配列。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-431">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line.</span></span> <span data-ttu-id="7d0dd-432">配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-432">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="7d0dd-433">配列。&lt;文字列&gt; | 配列。&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="7d0dd-433">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="7d0dd-434">電子メール アドレスを含む文字列の配列、または Cc 行の各受信者の `EmailAddressDetails` オブジェクトを含む配列。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-434">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line.</span></span> <span data-ttu-id="7d0dd-435">配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-435">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="7d0dd-436">配列。&lt;文字列&gt; | 配列。&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="7d0dd-436">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="7d0dd-437">電子メール アドレスを含む文字列の配列、または Bcc 行の各受信者の `EmailAddressDetails` オブジェクトを含む配列。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-437">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line.</span></span> <span data-ttu-id="7d0dd-438">配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-438">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="7d0dd-439">文字列</span><span class="sxs-lookup"><span data-stu-id="7d0dd-439">String</span></span> | <span data-ttu-id="7d0dd-440">メッセージの件名を含む文字列。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-440">A string containing the subject of the message.</span></span> <span data-ttu-id="7d0dd-441">文字列は最大 255 文字に制限されます。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-441">The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="7d0dd-442">文字列</span><span class="sxs-lookup"><span data-stu-id="7d0dd-442">String</span></span> | <span data-ttu-id="7d0dd-443">メッセージの HTML 本文。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-443">The HTML body of the message.</span></span> <span data-ttu-id="7d0dd-444">本文の内容は、最大サイズが 32 KB に制限されます。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-444">The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="7d0dd-445">配列。&lt;オブジェクト&gt;</span><span class="sxs-lookup"><span data-stu-id="7d0dd-445">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="7d0dd-446">添付ファイルまたは添付アイテムである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-446">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="7d0dd-447">文字列</span><span class="sxs-lookup"><span data-stu-id="7d0dd-447">String</span></span> | <span data-ttu-id="7d0dd-p129">添付ファイルの種類を示します。添付ファイルの場合は`file`、添付項目の場合は`item`でなければなりません。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-p129">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="7d0dd-450">文字列</span><span class="sxs-lookup"><span data-stu-id="7d0dd-450">String</span></span> | <span data-ttu-id="7d0dd-451">添付ファイル名を含む文字列で、255 文字以内で入力が可能です。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-451">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="7d0dd-452">文字列</span><span class="sxs-lookup"><span data-stu-id="7d0dd-452">String</span></span> | <span data-ttu-id="7d0dd-p130">`type`が`file`に設定されている場合にのみ使用されます。ファイルの場所の URIです。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-p130">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="7d0dd-455">ブール値</span><span class="sxs-lookup"><span data-stu-id="7d0dd-455">Boolean</span></span> | <span data-ttu-id="7d0dd-p131">`type`が`file`に設定されている場合にのみ使用されます。`true`の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-p131">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="7d0dd-458">文字列</span><span class="sxs-lookup"><span data-stu-id="7d0dd-458">String</span></span> | <span data-ttu-id="7d0dd-459">`type` が `item` に設定されている場合にのみ使用されます。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-459">Only used if `type` is set to `item`.</span></span> <span data-ttu-id="7d0dd-460">新しいメッセージに添付する、既存の電子メールの EWS 項目の id です。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-460">The EWS item id of the existing e-mail you want to attach to the new message.</span></span> <span data-ttu-id="7d0dd-461">最長 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-461">This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="7d0dd-462">要件</span><span class="sxs-lookup"><span data-stu-id="7d0dd-462">Requirements</span></span>

|<span data-ttu-id="7d0dd-463">必要条件</span><span class="sxs-lookup"><span data-stu-id="7d0dd-463">Requirement</span></span>| <span data-ttu-id="7d0dd-464">値</span><span class="sxs-lookup"><span data-stu-id="7d0dd-464">Value</span></span>|
|---|---|
|[<span data-ttu-id="7d0dd-465">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="7d0dd-465">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7d0dd-466">1.6</span><span class="sxs-lookup"><span data-stu-id="7d0dd-466">-16</span></span> |
|[<span data-ttu-id="7d0dd-467">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="7d0dd-467">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7d0dd-468">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7d0dd-468">ReadItem</span></span>|
|[<span data-ttu-id="7d0dd-469">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="7d0dd-469">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7d0dd-470">読み取り</span><span class="sxs-lookup"><span data-stu-id="7d0dd-470">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7d0dd-471">例</span><span class="sxs-lookup"><span data-stu-id="7d0dd-471">Example</span></span>

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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="7d0dd-472">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="7d0dd-472">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="7d0dd-473">REST API または Exchange Web サービスを呼び出すために使用するトークンを含む文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-473">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="7d0dd-p133">`getCallbackTokenAsync` メソッドは、非同期の呼び出しを行なって、ユーザーのメールボックスをホストする Exchange Server から opaque トークンを取得します。コールバック トークンの有効期間は 5 分です。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-p133">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="7d0dd-476">可能な場合は常に、アドインでは Exchange Web サービスの代わりに REST API を使用することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-476">Note: It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span> 

<span data-ttu-id="7d0dd-477">**REST トークン**</span><span class="sxs-lookup"><span data-stu-id="7d0dd-477">**REST Tokens**</span></span>

<span data-ttu-id="7d0dd-p134">REST トークンが要求された場合 (`options.isRest = true`) には、作成されたトークンは Exchange Web サービスの呼び出しを認証するためには機能しません。このトークンは、アドインがマニフェストで [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) アクセス許可を指定しない限り、現在の項目およびその添付ファイルへの読み取り専用の範囲に制限されます。`ReadWriteMailbox` アクセス許可が指定された場合には、作成されるトークンは、メールを送信する機能など、メール、予定表、連絡先への読み取り/書き込みアクセスを付与します。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-p134">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="7d0dd-481">アドインでは、`restUrl`プロパティを使用して、REST API 呼び出しを行うときに使用する正しい URL を決定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-481">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="7d0dd-482">**EWS トークン**</span><span class="sxs-lookup"><span data-stu-id="7d0dd-482">**EWS Tokens**</span></span>

<span data-ttu-id="7d0dd-p135">EWS トークンが要求された場合(`options.isRest = false`) には、作成されるトークンは REST API の呼び出しを認証するためには機能しません。このトークンは、現在の項目にアクセスできる範囲に制限されます。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-p135">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="7d0dd-485">アドインでは、`ewsUrl` プロパティを使用して、EWS 呼び出しを行うときに使用する正しい URL を決定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-485">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7d0dd-486">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="7d0dd-486">Parameters:</span></span>

|<span data-ttu-id="7d0dd-487">名前</span><span class="sxs-lookup"><span data-stu-id="7d0dd-487">Name</span></span>| <span data-ttu-id="7d0dd-488">種類</span><span class="sxs-lookup"><span data-stu-id="7d0dd-488">Type</span></span>| <span data-ttu-id="7d0dd-489">属性</span><span class="sxs-lookup"><span data-stu-id="7d0dd-489">Attributes</span></span>| <span data-ttu-id="7d0dd-490">説明</span><span class="sxs-lookup"><span data-stu-id="7d0dd-490">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="7d0dd-491">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="7d0dd-491">Object</span></span> | <span data-ttu-id="7d0dd-492">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="7d0dd-492">&lt;optional&gt;</span></span> | <span data-ttu-id="7d0dd-493">次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-493">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="7d0dd-494">ブール値</span><span class="sxs-lookup"><span data-stu-id="7d0dd-494">Boolean</span></span> |  <span data-ttu-id="7d0dd-495">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="7d0dd-495">&lt;optional&gt;</span></span> | <span data-ttu-id="7d0dd-p136">提供されたトークンを Outlook REST API または Exchange Web サービスに使用するかどうかを決定します。既定値は、`false`です。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-p136">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="7d0dd-498">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="7d0dd-498">Object</span></span> |  <span data-ttu-id="7d0dd-499">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="7d0dd-499">&lt;optional&gt;</span></span> | <span data-ttu-id="7d0dd-500">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-500">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="7d0dd-501">関数</span><span class="sxs-lookup"><span data-stu-id="7d0dd-501">function</span></span>||<span data-ttu-id="7d0dd-p137">メソッドが完了すると、`callback` パラメータに渡された関数が、単一パラメータ `asyncResult` で呼び出されます。このパラメータは、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトです。トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-p137">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7d0dd-504">要件</span><span class="sxs-lookup"><span data-stu-id="7d0dd-504">Requirements</span></span>

|<span data-ttu-id="7d0dd-505">必要条件</span><span class="sxs-lookup"><span data-stu-id="7d0dd-505">Requirement</span></span>| <span data-ttu-id="7d0dd-506">値</span><span class="sxs-lookup"><span data-stu-id="7d0dd-506">Value</span></span>|
|---|---|
|[<span data-ttu-id="7d0dd-507">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="7d0dd-507">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7d0dd-508">1.5</span><span class="sxs-lookup"><span data-stu-id="7d0dd-508">1.5</span></span> |
|[<span data-ttu-id="7d0dd-509">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="7d0dd-509">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7d0dd-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7d0dd-510">ReadItem</span></span>|
|[<span data-ttu-id="7d0dd-511">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="7d0dd-511">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7d0dd-512">新規作成と閲覧</span><span class="sxs-lookup"><span data-stu-id="7d0dd-512">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="7d0dd-513">例</span><span class="sxs-lookup"><span data-stu-id="7d0dd-513">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="7d0dd-514">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="7d0dd-514">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="7d0dd-515">Exchange Server から添付ファイルやアイテムを取得するために使用するトークンを含む文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-515">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="7d0dd-p138">`getCallbackTokenAsync` メソッドは、非同期の呼び出しを行なって、ユーザーのメールボックスをホストする Exchange Server から opaque トークンを取得します。コールバック トークンの有効期間は 5 分です。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-p138">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="7d0dd-p139">このトークンと、添付ファイル識別子または項目識別子は、サードパーティーのシステムに渡すことができます。サードパーティーのシステムでは、添付ファイルまたは項目を返すための Exchange Web サービス (EWS) の [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) または [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) 操作を呼び出すのに、トークンをベアラー承認トークンとして使用します。たとえば、リモート サービスを作成して[選択した項目から添付ファイルを取得](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item)することができます。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-p139">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="7d0dd-521">アプリでは、閲覧モードで `getCallbackTokenAsync` メソッドを呼び出すために、 **ReadItem** アクセス許可をアプリのマニフェストで指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-521">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="7d0dd-p140">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出して、`getCallbackTokenAsync` メソッドに渡すための項目識別子を取得する必要があります。アプリには、`saveAsync` メソッドを呼び出すために **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-p140">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7d0dd-524">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="7d0dd-524">Parameters:</span></span>

|<span data-ttu-id="7d0dd-525">名前</span><span class="sxs-lookup"><span data-stu-id="7d0dd-525">Name</span></span>| <span data-ttu-id="7d0dd-526">種類</span><span class="sxs-lookup"><span data-stu-id="7d0dd-526">Type</span></span>| <span data-ttu-id="7d0dd-527">属性</span><span class="sxs-lookup"><span data-stu-id="7d0dd-527">Attributes</span></span>| <span data-ttu-id="7d0dd-528">説明</span><span class="sxs-lookup"><span data-stu-id="7d0dd-528">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="7d0dd-529">関数</span><span class="sxs-lookup"><span data-stu-id="7d0dd-529">function</span></span>||<span data-ttu-id="7d0dd-p141">メソッドが完了すると、`callback` パラメータに渡された関数が、単一パラメータ `asyncResult` で呼び出されます。このパラメータは、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトです。トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-p141">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="7d0dd-532">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="7d0dd-532">Object</span></span>| <span data-ttu-id="7d0dd-533">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="7d0dd-533">&lt;optional&gt;</span></span>|<span data-ttu-id="7d0dd-534">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-534">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7d0dd-535">要件</span><span class="sxs-lookup"><span data-stu-id="7d0dd-535">Requirements</span></span>

|<span data-ttu-id="7d0dd-536">必要条件</span><span class="sxs-lookup"><span data-stu-id="7d0dd-536">Requirement</span></span>| <span data-ttu-id="7d0dd-537">値</span><span class="sxs-lookup"><span data-stu-id="7d0dd-537">Value</span></span>|
|---|---|
|[<span data-ttu-id="7d0dd-538">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="7d0dd-538">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7d0dd-539">1.3</span><span class="sxs-lookup"><span data-stu-id="7d0dd-539">1.3</span></span>|
|[<span data-ttu-id="7d0dd-540">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="7d0dd-540">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7d0dd-541">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7d0dd-541">ReadItem</span></span>|
|[<span data-ttu-id="7d0dd-542">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="7d0dd-542">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7d0dd-543">新規作成と閲覧</span><span class="sxs-lookup"><span data-stu-id="7d0dd-543">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="7d0dd-544">例</span><span class="sxs-lookup"><span data-stu-id="7d0dd-544">Example</span></span>

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="7d0dd-545">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="7d0dd-545">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="7d0dd-546">ユーザーと Office アドインを識別するトークンを取得します。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-546">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="7d0dd-547">`getUserIdentityTokenAsync` メソッドは、[アドインとユーザーをサードパーティのシステムで識別して認証](https://docs.microsoft.com/outlook/add-ins/authentication)するのに使用できるトークンを返します。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-547">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](https://docs.microsoft.com/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="7d0dd-548">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="7d0dd-548">Parameters:</span></span>

|<span data-ttu-id="7d0dd-549">名前</span><span class="sxs-lookup"><span data-stu-id="7d0dd-549">Name</span></span>| <span data-ttu-id="7d0dd-550">種類</span><span class="sxs-lookup"><span data-stu-id="7d0dd-550">Type</span></span>| <span data-ttu-id="7d0dd-551">属性</span><span class="sxs-lookup"><span data-stu-id="7d0dd-551">Attributes</span></span>| <span data-ttu-id="7d0dd-552">説明</span><span class="sxs-lookup"><span data-stu-id="7d0dd-552">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="7d0dd-553">関数</span><span class="sxs-lookup"><span data-stu-id="7d0dd-553">function</span></span>||<span data-ttu-id="7d0dd-554">メソッドが完了すると、`callback` パラメータに渡された関数が、単一パラメータ `asyncResult` で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-554">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="7d0dd-555">トークンは、`asyncResult.value`プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-555">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="7d0dd-556">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="7d0dd-556">Object</span></span>| <span data-ttu-id="7d0dd-557">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="7d0dd-557">&lt;optional&gt;</span></span>|<span data-ttu-id="7d0dd-558">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-558">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7d0dd-559">要件</span><span class="sxs-lookup"><span data-stu-id="7d0dd-559">Requirements</span></span>

|<span data-ttu-id="7d0dd-560">必要条件</span><span class="sxs-lookup"><span data-stu-id="7d0dd-560">Requirement</span></span>| <span data-ttu-id="7d0dd-561">値</span><span class="sxs-lookup"><span data-stu-id="7d0dd-561">Value</span></span>|
|---|---|
|[<span data-ttu-id="7d0dd-562">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="7d0dd-562">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7d0dd-563">1.0以降</span><span class="sxs-lookup"><span data-stu-id="7d0dd-563">1.0</span></span>|
|[<span data-ttu-id="7d0dd-564">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="7d0dd-564">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7d0dd-565">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7d0dd-565">ReadItem</span></span>|
|[<span data-ttu-id="7d0dd-566">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="7d0dd-566">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7d0dd-567">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="7d0dd-567">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="7d0dd-568">例</span><span class="sxs-lookup"><span data-stu-id="7d0dd-568">Example</span></span>

```js
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="7d0dd-569">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="7d0dd-569">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="7d0dd-570">ユーザーのメールボックスをホストしている Exchange Server上の Exchange Web サービス (EWS) のサービスに対して非同期の要求を行います。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-570">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="7d0dd-571">このメソッドは、次のシナリオではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-571">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="7d0dd-572">Outlook for iOS または Outlook for Android で</span><span class="sxs-lookup"><span data-stu-id="7d0dd-572">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="7d0dd-573">アドインが Gmail のメールボックスにロードされる場合</span><span class="sxs-lookup"><span data-stu-id="7d0dd-573">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="7d0dd-574">これらの場合、アドインではユーザーのメールボックスにアクセスするために、代わりに [REST API を使用する](https://docs.microsoft.com/outlook/add-ins/use-rest-api)必要があります。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-574">In these cases, add-ins should [use REST APIs](https://docs.microsoft.com/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="7d0dd-575">`makeEwsRequestAsync` メソッドは、アドインの代わりに Exchange に EWS 要求を送信します。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-575">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="7d0dd-576">サポートされている EWS 操作の一覧については、 「[ Outlook アドインから Web サービスを呼び出す](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-576">See [Call web services from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="7d0dd-577">`makeEwsRequestAsync` メソッドで、フォルダー関連アイテムを要求することはできません。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-577">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="7d0dd-578">XML 要求では、UTF-8 エンコードを指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-578">The XML request must specify UTF-8 encoding.</span></span>

```
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="7d0dd-p143">アドインには、`makeEwsRequestAsync` メソッドを使用するために **ReadWriteMailbox** アクセス許可が必要です。**ReadWriteMailbox** アクセス許可と、`makeEwsRequestAsync` メソッドで呼び出すことのできる EWS 操作の使用の詳細については、「[ユーザーのメールボックスへのメール アドイン アクセスのアクセス許可を指定する](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-p143">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="7d0dd-581">サーバー管理者は、クライアント アクセス サーバー の EWS ディレクトリ上で `OAuthAuthentication` を true に設定して、`makeEwsRequestAsync` メソッドで EWS 要求を行えるようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-581">NOTE: The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="7d0dd-582">バージョンの相違点</span><span class="sxs-lookup"><span data-stu-id="7d0dd-582">Version differences</span></span>

<span data-ttu-id="7d0dd-583">バージョン 15.0.4535.1004 より前のバージョンの Outlook で実行しているメール アプリで `makeEwsRequestAsync` メソッドを使用する場合には、エンコード値を `ISO-8859-1` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-583">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="7d0dd-p144">メール アプリが Outlook on the web で実行されている場合には、エンコード値を設定する必要はありません。メールボックスを使用してメール アプリが Outlook で実行されているのか、Outlook on the web で実行されているのかを判断する必要があります。mailbox.diagnostics.hostVersion プロパティを使用すれば、どのバージョンの Outlook が実行されているのかがわかります。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-p144">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7d0dd-587">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="7d0dd-587">Parameters:</span></span>

|<span data-ttu-id="7d0dd-588">名前</span><span class="sxs-lookup"><span data-stu-id="7d0dd-588">Name</span></span>| <span data-ttu-id="7d0dd-589">種類</span><span class="sxs-lookup"><span data-stu-id="7d0dd-589">Type</span></span>| <span data-ttu-id="7d0dd-590">属性</span><span class="sxs-lookup"><span data-stu-id="7d0dd-590">Attributes</span></span>| <span data-ttu-id="7d0dd-591">説明</span><span class="sxs-lookup"><span data-stu-id="7d0dd-591">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="7d0dd-592">文字列</span><span class="sxs-lookup"><span data-stu-id="7d0dd-592">String</span></span>||<span data-ttu-id="7d0dd-593">EWS 要求。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-593">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="7d0dd-594">関数</span><span class="sxs-lookup"><span data-stu-id="7d0dd-594">function</span></span>||<span data-ttu-id="7d0dd-595">メソッドが完了すると、`callback` パラメータで渡された関数が、単一パラメータ `asyncResult` で呼び出されます。このパラメータは、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-595">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="7d0dd-596">EWS 呼び出しの XML 結果は、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-596">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="7d0dd-597">結果のサイズが 1 MB を超えている場合、エラー メッセージが返されます。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-597">If the result exceeds 1 MB in size, an error message is returned instead.| | Object| optional|Any state data that is passed to the asynchronous method.|</span></span>|
|`userContext`| <span data-ttu-id="7d0dd-598">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="7d0dd-598">Object</span></span>| <span data-ttu-id="7d0dd-599">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="7d0dd-599">&lt;optional&gt;</span></span>|<span data-ttu-id="7d0dd-600">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-600">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7d0dd-601">要件</span><span class="sxs-lookup"><span data-stu-id="7d0dd-601">Requirements</span></span>

|<span data-ttu-id="7d0dd-602">必要条件</span><span class="sxs-lookup"><span data-stu-id="7d0dd-602">Requirement</span></span>| <span data-ttu-id="7d0dd-603">値</span><span class="sxs-lookup"><span data-stu-id="7d0dd-603">Value</span></span>|
|---|---|
|[<span data-ttu-id="7d0dd-604">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="7d0dd-604">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7d0dd-605">1.0以降</span><span class="sxs-lookup"><span data-stu-id="7d0dd-605">1.0</span></span>|
|[<span data-ttu-id="7d0dd-606">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="7d0dd-606">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7d0dd-607">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="7d0dd-607">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="7d0dd-608">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="7d0dd-608">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7d0dd-609">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="7d0dd-609">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="7d0dd-610">例</span><span class="sxs-lookup"><span data-stu-id="7d0dd-610">Example</span></span>

<span data-ttu-id="7d0dd-611">次の例は、`makeEwsRequestAsync` を呼び出し、`GetItem` 操作を使って項目の件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="7d0dd-611">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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