
# <a name="mailbox"></a><span data-ttu-id="d4771-101">mailbox</span><span class="sxs-lookup"><span data-stu-id="d4771-101">mailbox</span></span>

### <span data-ttu-id="d4771-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span><span class="sxs-lookup"><span data-stu-id="d4771-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span></span>

<span data-ttu-id="d4771-104">Microsoft Outlook と Microsoft Outlook on the web の Outlook アドイン オブジェクト モデルへのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="d4771-104">Provides access to the Outlook Add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d4771-105">要件</span><span class="sxs-lookup"><span data-stu-id="d4771-105">Requirements</span></span>

|<span data-ttu-id="d4771-106">要件</span><span class="sxs-lookup"><span data-stu-id="d4771-106">Requirement</span></span>| <span data-ttu-id="d4771-107">値</span><span class="sxs-lookup"><span data-stu-id="d4771-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="d4771-108">最小限のメールボックス要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d4771-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d4771-109">1.0</span><span class="sxs-lookup"><span data-stu-id="d4771-109">1.0</span></span>|
|[<span data-ttu-id="d4771-110">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d4771-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d4771-111">制限あり</span><span class="sxs-lookup"><span data-stu-id="d4771-111">Restricted</span></span>|
|[<span data-ttu-id="d4771-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d4771-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d4771-113">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="d4771-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="d4771-114">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="d4771-114">Members and methods</span></span>

| <span data-ttu-id="d4771-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="d4771-115">Member</span></span> | <span data-ttu-id="d4771-116">種類</span><span class="sxs-lookup"><span data-stu-id="d4771-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="d4771-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="d4771-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="d4771-118">メンバー</span><span class="sxs-lookup"><span data-stu-id="d4771-118">Member</span></span> |
| [<span data-ttu-id="d4771-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="d4771-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="d4771-120">メンバー</span><span class="sxs-lookup"><span data-stu-id="d4771-120">Member</span></span> |
| [<span data-ttu-id="d4771-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="d4771-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="d4771-122">メソッド</span><span class="sxs-lookup"><span data-stu-id="d4771-122">Method</span></span> |
| [<span data-ttu-id="d4771-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="d4771-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="d4771-124">メソッド</span><span class="sxs-lookup"><span data-stu-id="d4771-124">Method</span></span> |
| [<span data-ttu-id="d4771-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="d4771-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) | <span data-ttu-id="d4771-126">メソッド</span><span class="sxs-lookup"><span data-stu-id="d4771-126">Method</span></span> |
| [<span data-ttu-id="d4771-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="d4771-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="d4771-128">メソッド</span><span class="sxs-lookup"><span data-stu-id="d4771-128">Method</span></span> |
| [<span data-ttu-id="d4771-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="d4771-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="d4771-130">メソッド</span><span class="sxs-lookup"><span data-stu-id="d4771-130">Method</span></span> |
| [<span data-ttu-id="d4771-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="d4771-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="d4771-132">メソッド</span><span class="sxs-lookup"><span data-stu-id="d4771-132">Method</span></span> |
| [<span data-ttu-id="d4771-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="d4771-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="d4771-134">メソッド</span><span class="sxs-lookup"><span data-stu-id="d4771-134">Method</span></span> |
| [<span data-ttu-id="d4771-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="d4771-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="d4771-136">メソッド</span><span class="sxs-lookup"><span data-stu-id="d4771-136">Method</span></span> |
| [<span data-ttu-id="d4771-137">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="d4771-137">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="d4771-138">メソッド</span><span class="sxs-lookup"><span data-stu-id="d4771-138">Method</span></span> |
| [<span data-ttu-id="d4771-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="d4771-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="d4771-140">メソッド</span><span class="sxs-lookup"><span data-stu-id="d4771-140">Method</span></span> |
| [<span data-ttu-id="d4771-141">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="d4771-141">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="d4771-142">メソッド</span><span class="sxs-lookup"><span data-stu-id="d4771-142">Method</span></span> |
| [<span data-ttu-id="d4771-143">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="d4771-143">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="d4771-144">メソッド</span><span class="sxs-lookup"><span data-stu-id="d4771-144">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="d4771-145">名前空間</span><span class="sxs-lookup"><span data-stu-id="d4771-145">Namespaces</span></span>

<span data-ttu-id="d4771-146">[diagnostics](Office.context.mailbox.diagnostics.md): Outlook アドインに診断情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="d4771-146">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="d4771-147">[item](Office.context.mailbox.item.md):Outlook アドインのメッセージや予定にアクセスするためのメソッドとプロパティを提供します。</span><span class="sxs-lookup"><span data-stu-id="d4771-147">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="d4771-148">[userProfile](Office.context.mailbox.userProfile.md):Outlook アドインのユーザーに関する情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="d4771-148">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="d4771-149">メンバー</span><span class="sxs-lookup"><span data-stu-id="d4771-149">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="d4771-150">ewsUrl: 文字列</span><span class="sxs-lookup"><span data-stu-id="d4771-150">ewsUrl :String</span></span>

<span data-ttu-id="d4771-p102">このメール アカウントの Exchange Web サービス (EWS) エンドポイントの URL を取得します。閲覧モードのみです。</span><span class="sxs-lookup"><span data-stu-id="d4771-p102">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="d4771-153">このメンバーは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="d4771-153">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d4771-p103">`ewsUrl` 値は、リモート サービスで、ユーザーのメールボックスに EWS 呼び出しを行うために使用することができます。たとえば、[選択した項目から添付ファイルを取得する](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="d4771-p103">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="d4771-156">閲覧モードで \*\*  メンバーを呼び出すには、アプリでマニフェスト内に \*\*  ReadItem`ewsUrl`    アクセス許可を指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="d4771-156">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="d4771-p104">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback)メソッドを呼び出してから、`ewsUrl`メンバーを使用する必要があります。アプリには、\*\* メソッドを呼び出す \*\* ReadWriteItem`saveAsync`   アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="d4771-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="d4771-159">種類:</span><span class="sxs-lookup"><span data-stu-id="d4771-159">Type:</span></span>

*   <span data-ttu-id="d4771-160">文字列</span><span class="sxs-lookup"><span data-stu-id="d4771-160">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d4771-161">要件</span><span class="sxs-lookup"><span data-stu-id="d4771-161">Requirements</span></span>

|<span data-ttu-id="d4771-162">要件</span><span class="sxs-lookup"><span data-stu-id="d4771-162">Requirement</span></span>| <span data-ttu-id="d4771-163">値</span><span class="sxs-lookup"><span data-stu-id="d4771-163">Value</span></span>|
|---|---|
|[<span data-ttu-id="d4771-164">最小限のメールボックス要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d4771-164">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d4771-165">1.0</span><span class="sxs-lookup"><span data-stu-id="d4771-165">1.0</span></span>|
|[<span data-ttu-id="d4771-166">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d4771-166">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d4771-167">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d4771-167">ReadItem</span></span>|
|[<span data-ttu-id="d4771-168">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d4771-168">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d4771-169">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="d4771-169">Compose or read</span></span>|

#### <a name="resturl-string"></a><span data-ttu-id="d4771-170">restUrl: 文字列</span><span class="sxs-lookup"><span data-stu-id="d4771-170">restUrl :String</span></span>

<span data-ttu-id="d4771-171">この電子メール アカウントの REST エンドポイントの URL を取得します。</span><span class="sxs-lookup"><span data-stu-id="d4771-171">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="d4771-172">`restUrl` 値は、ユーザーのメールボックスに [REST API](https://docs.microsoft.com/outlook/rest/) 呼び出しを行うために使用することができます。</span><span class="sxs-lookup"><span data-stu-id="d4771-172">The `restUrl` value can be used to make [REST API](https://docs.microsoft.com/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="d4771-173">閲覧モードで \*\*  メンバーを呼び出すには、アプリでマニフェスト内に \*\*  ReadItem`restUrl`    アクセス許可を指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="d4771-173">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="d4771-p105">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback)メソッドを呼び出してから、`restUrl`メンバーを使用する必要があります。アプリには、\*\* メソッドを呼び出す \*\* ReadWriteItem`saveAsync`   アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="d4771-p105">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

> [!NOTE]
> <span data-ttu-id="d4771-176">カスタム REST URL が構成された Exchange 2016 またはそれ以降のオンプレミスのインストールに接続されている Outlook クライアントは、`restUrl` に無効な値を返します。</span><span class="sxs-lookup"><span data-stu-id="d4771-176">Note: Outlook clients connected to on-premises installations of Exchange 2016 with a custom REST URL configured will return an invalid value for `restUrl`.</span></span>

##### <a name="type"></a><span data-ttu-id="d4771-177">種類:</span><span class="sxs-lookup"><span data-stu-id="d4771-177">Type:</span></span>

*   <span data-ttu-id="d4771-178">文字列</span><span class="sxs-lookup"><span data-stu-id="d4771-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d4771-179">要件</span><span class="sxs-lookup"><span data-stu-id="d4771-179">Requirements</span></span>

|<span data-ttu-id="d4771-180">要件</span><span class="sxs-lookup"><span data-stu-id="d4771-180">Requirement</span></span>| <span data-ttu-id="d4771-181">値</span><span class="sxs-lookup"><span data-stu-id="d4771-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="d4771-182">最小限のメールボックス要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d4771-182">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d4771-183">1.5</span><span class="sxs-lookup"><span data-stu-id="d4771-183">1.5</span></span> |
|[<span data-ttu-id="d4771-184">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d4771-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d4771-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d4771-185">ReadItem</span></span>|
|[<span data-ttu-id="d4771-186">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d4771-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d4771-187">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="d4771-187">Compose or read</span></span>|

### <a name="methods"></a><span data-ttu-id="d4771-188">メソッド</span><span class="sxs-lookup"><span data-stu-id="d4771-188">Methods</span></span>

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="d4771-189">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="d4771-189">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="d4771-190">サポートされているイベントのイベント ハンドラを追加します。</span><span class="sxs-lookup"><span data-stu-id="d4771-190">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="d4771-p106">現在サポートされている唯一のイベントの種類は `Office.EventType.ItemChanged` です。これはユーザーが新しいアイテムを選択したときに呼び出されます。このイベントは、ピン留め可能な作業ウィンドウを実装するアドインで使用され、現在選択されているアイテムに基づいて作業ウィンドウ UI をアドインで更新できるようにします。</span><span class="sxs-lookup"><span data-stu-id="d4771-p106">Currently the only supported event type is `Office.EventType.ItemChanged`, which is invoked when the user selects a new item. This event is used by add-ins that implement a pinnable taskpane, and allows the add-in to refresh the taskpane UI based on the currently selected item.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d4771-193">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="d4771-193">Parameters:</span></span>

| <span data-ttu-id="d4771-194">名前</span><span class="sxs-lookup"><span data-stu-id="d4771-194">Name</span></span> | <span data-ttu-id="d4771-195">種類</span><span class="sxs-lookup"><span data-stu-id="d4771-195">Type</span></span> | <span data-ttu-id="d4771-196">属性</span><span class="sxs-lookup"><span data-stu-id="d4771-196">Attributes</span></span> | <span data-ttu-id="d4771-197">説明</span><span class="sxs-lookup"><span data-stu-id="d4771-197">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="d4771-198">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="d4771-198">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="d4771-199">ハンドラを呼び出す必要のあるイベント。</span><span class="sxs-lookup"><span data-stu-id="d4771-199">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="d4771-200">関数</span><span class="sxs-lookup"><span data-stu-id="d4771-200">Function</span></span> || <span data-ttu-id="d4771-p107">イベントを処理する関数。この関数は、オブジェクト リテラルである単一パラメータを受け入れる必要があります。パラメータの `type` プロパティは、`eventType`  に渡される `addHandlerAsync`  パラメータと一致します。</span><span class="sxs-lookup"><span data-stu-id="d4771-p107">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="d4771-204">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d4771-204">Object</span></span> | <span data-ttu-id="d4771-205">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="d4771-205">&lt;optional&gt;</span></span> | <span data-ttu-id="d4771-206">以下のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="d4771-206">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="d4771-207">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d4771-207">Object</span></span> | <span data-ttu-id="d4771-208">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="d4771-208">&lt;optional&gt;</span></span> | <span data-ttu-id="d4771-209">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="d4771-209">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="d4771-210">関数</span><span class="sxs-lookup"><span data-stu-id="d4771-210">function</span></span>| <span data-ttu-id="d4771-211">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="d4771-211">&lt;optional&gt;</span></span>|<span data-ttu-id="d4771-212">メソッドが完了すると、`callback` パラメータで渡された関数が、単一パラメータ `asyncResult` で呼び出されます。このパラメータは、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="d4771-212">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d4771-213">要件</span><span class="sxs-lookup"><span data-stu-id="d4771-213">Requirements</span></span>

|<span data-ttu-id="d4771-214">要件</span><span class="sxs-lookup"><span data-stu-id="d4771-214">Requirement</span></span>| <span data-ttu-id="d4771-215">値</span><span class="sxs-lookup"><span data-stu-id="d4771-215">Value</span></span>|
|---|---|
|[<span data-ttu-id="d4771-216">最小限のメールボックス要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d4771-216">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d4771-217">1.5</span><span class="sxs-lookup"><span data-stu-id="d4771-217">1.5</span></span> |
|[<span data-ttu-id="d4771-218">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d4771-218">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d4771-219">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d4771-219">ReadItem</span></span> |
|[<span data-ttu-id="d4771-220">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d4771-220">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d4771-221">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="d4771-221">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d4771-222">例</span><span class="sxs-lookup"><span data-stu-id="d4771-222">Example</span></span>

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

####  <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="d4771-223">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="d4771-223">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="d4771-224">REST 用に書式設定された項目 ID を EWS 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="d4771-224">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="d4771-225">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="d4771-225">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d4771-p108">REST API ([Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) や [Microsoft Graph](http://graph.microsoft.io/) など) 経由で取得された項目 ID は、Exchange Web サービス (EWS) で使用される形式とは異なる形式を使用します。`convertToEwsId` メソッドは、REST 形式の ID を EWS 用の適切な形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="d4771-p108">Item IDs retrieved via a REST API (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](http://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d4771-228">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="d4771-228">Parameters:</span></span>

|<span data-ttu-id="d4771-229">名前</span><span class="sxs-lookup"><span data-stu-id="d4771-229">Name</span></span>| <span data-ttu-id="d4771-230">種類</span><span class="sxs-lookup"><span data-stu-id="d4771-230">Type</span></span>| <span data-ttu-id="d4771-231">説明</span><span class="sxs-lookup"><span data-stu-id="d4771-231">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="d4771-232">文字列</span><span class="sxs-lookup"><span data-stu-id="d4771-232">String</span></span>|<span data-ttu-id="d4771-233">Outlook REST API 用に書式設定された項目 ID</span><span class="sxs-lookup"><span data-stu-id="d4771-233">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="d4771-234">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="d4771-234">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.restversion)|<span data-ttu-id="d4771-235">項目 ID の取得に使用された Outlook REST API のバージョンを示す値。</span><span class="sxs-lookup"><span data-stu-id="d4771-235">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d4771-236">要件</span><span class="sxs-lookup"><span data-stu-id="d4771-236">Requirements</span></span>

|<span data-ttu-id="d4771-237">要件</span><span class="sxs-lookup"><span data-stu-id="d4771-237">Requirement</span></span>| <span data-ttu-id="d4771-238">値</span><span class="sxs-lookup"><span data-stu-id="d4771-238">Value</span></span>|
|---|---|
|[<span data-ttu-id="d4771-239">最小限のメールボックス要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d4771-239">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d4771-240">1.3</span><span class="sxs-lookup"><span data-stu-id="d4771-240">1.3</span></span>|
|[<span data-ttu-id="d4771-241">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d4771-241">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d4771-242">制限あり</span><span class="sxs-lookup"><span data-stu-id="d4771-242">Restricted</span></span>|
|[<span data-ttu-id="d4771-243">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d4771-243">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d4771-244">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="d4771-244">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d4771-245">戻り値 :</span><span class="sxs-lookup"><span data-stu-id="d4771-245">Returns:</span></span>

<span data-ttu-id="d4771-246">種類: 文字列</span><span class="sxs-lookup"><span data-stu-id="d4771-246">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="d4771-247">例</span><span class="sxs-lookup"><span data-stu-id="d4771-247">Example</span></span>

```
// Get an item's ID from a REST API
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the
// Outlook Mail API
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime"></a><span data-ttu-id="d4771-248">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_5/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="d4771-248">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_5/office.LocalClientTime)}</span></span>

<span data-ttu-id="d4771-249">クライアントの現地時間の時間情報が含まれている辞書を取得します。</span><span class="sxs-lookup"><span data-stu-id="d4771-249">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="d4771-p109">Outlook 用メール アプリや Outlook Web App で使う日付と時刻は、異なるタイム ゾーンを使うことができます。Outlook では、クライアント コンピューターのタイム ゾーンを使います。Outlook Web App では、Exchange 管理センター (EAC) で設定されたタイム ゾーンを使います。ユーザー インターフェイスに表示される値が、常にユーザーが期待するタイム ゾーンと一致するように日付と時刻の値を処理する必要があります。</span><span class="sxs-lookup"><span data-stu-id="d4771-p109">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="d4771-p110">Outlook でメール アプリが実行されている場合、`convertToLocalClientTime` メソッドは、クライアント コンピューターのタイム ゾーンに設定された値で辞書オブジェクトを返します。Outlook Web Apps でメール アプリが実行されている場合、`convertToLocalClientTime` メソッドは、EAC で指定されたタイム ゾーンに設定された値で辞書オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="d4771-p110">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d4771-255">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="d4771-255">Parameters:</span></span>

|<span data-ttu-id="d4771-256">名前</span><span class="sxs-lookup"><span data-stu-id="d4771-256">Name</span></span>| <span data-ttu-id="d4771-257">種類</span><span class="sxs-lookup"><span data-stu-id="d4771-257">Type</span></span>| <span data-ttu-id="d4771-258">説明</span><span class="sxs-lookup"><span data-stu-id="d4771-258">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="d4771-259">Date</span><span class="sxs-lookup"><span data-stu-id="d4771-259">Date</span></span>|<span data-ttu-id="d4771-260">Date オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d4771-260">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d4771-261">要件</span><span class="sxs-lookup"><span data-stu-id="d4771-261">Requirements</span></span>

|<span data-ttu-id="d4771-262">要件</span><span class="sxs-lookup"><span data-stu-id="d4771-262">Requirement</span></span>| <span data-ttu-id="d4771-263">値</span><span class="sxs-lookup"><span data-stu-id="d4771-263">Value</span></span>|
|---|---|
|[<span data-ttu-id="d4771-264">最小限のメールボックス要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d4771-264">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d4771-265">1.0</span><span class="sxs-lookup"><span data-stu-id="d4771-265">1.0</span></span>|
|[<span data-ttu-id="d4771-266">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d4771-266">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d4771-267">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d4771-267">ReadItem</span></span>|
|[<span data-ttu-id="d4771-268">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d4771-268">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d4771-269">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="d4771-269">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d4771-270">戻り値 :</span><span class="sxs-lookup"><span data-stu-id="d4771-270">Returns:</span></span>

<span data-ttu-id="d4771-271">種類:[LocalClientTime](/javascript/api/outlook_1_5/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="d4771-271">Type: [LocalClientTime](/javascript/api/outlook_1_5/office.LocalClientTime)</span></span>

####  <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="d4771-272">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="d4771-272">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="d4771-273">EWS 用に書式設定された項目 ID を REST 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="d4771-273">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="d4771-274">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="d4771-274">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d4771-p111">EWS 経由または `itemId` プロパティ経由で取得される項目 ID では、REST API ([Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) または [Microsoft Graph](http://graph.microsoft.io/) など) で使用される形式とは異なる形式を使用します。`convertToRestId` メソッドは、EWS 形式の ID を REST 用の適切な形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="d4771-p111">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](http://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d4771-277">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="d4771-277">Parameters:</span></span>

|<span data-ttu-id="d4771-278">名前</span><span class="sxs-lookup"><span data-stu-id="d4771-278">Name</span></span>| <span data-ttu-id="d4771-279">種類</span><span class="sxs-lookup"><span data-stu-id="d4771-279">Type</span></span>| <span data-ttu-id="d4771-280">説明</span><span class="sxs-lookup"><span data-stu-id="d4771-280">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="d4771-281">文字列</span><span class="sxs-lookup"><span data-stu-id="d4771-281">String</span></span>|<span data-ttu-id="d4771-282">Exchange Web サービス (EWS) 用に書式設定された項目 ID</span><span class="sxs-lookup"><span data-stu-id="d4771-282">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="d4771-283">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="d4771-283">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.restversion)|<span data-ttu-id="d4771-284">変換後の ID とともに使用される Outlook REST API のバージョンを示す値。</span><span class="sxs-lookup"><span data-stu-id="d4771-284">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d4771-285">要件</span><span class="sxs-lookup"><span data-stu-id="d4771-285">Requirements</span></span>

|<span data-ttu-id="d4771-286">要件</span><span class="sxs-lookup"><span data-stu-id="d4771-286">Requirement</span></span>| <span data-ttu-id="d4771-287">値</span><span class="sxs-lookup"><span data-stu-id="d4771-287">Value</span></span>|
|---|---|
|[<span data-ttu-id="d4771-288">最小限のメールボックス要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d4771-288">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d4771-289">1.3</span><span class="sxs-lookup"><span data-stu-id="d4771-289">1.3</span></span>|
|[<span data-ttu-id="d4771-290">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d4771-290">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d4771-291">制限あり</span><span class="sxs-lookup"><span data-stu-id="d4771-291">Restricted</span></span>|
|[<span data-ttu-id="d4771-292">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d4771-292">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d4771-293">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="d4771-293">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d4771-294">戻り値 :</span><span class="sxs-lookup"><span data-stu-id="d4771-294">Returns:</span></span>

<span data-ttu-id="d4771-295">種類: 文字列</span><span class="sxs-lookup"><span data-stu-id="d4771-295">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="d4771-296">例</span><span class="sxs-lookup"><span data-stu-id="d4771-296">Example</span></span>

```
// Get the currently selected item's ID
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the
// Outlook Mail API
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="d4771-297">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="d4771-297">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="d4771-298">時間情報が含まれている辞書から Date オブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="d4771-298">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="d4771-299">`convertToUtcClientTime` メソッドは、ローカルの日付と時刻が含まれる辞書を、ローカルの日付と時刻の正しい値をもつ Date オブジェクトに変換します。</span><span class="sxs-lookup"><span data-stu-id="d4771-299">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d4771-300">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="d4771-300">Parameters:</span></span>

|<span data-ttu-id="d4771-301">名前</span><span class="sxs-lookup"><span data-stu-id="d4771-301">Name</span></span>| <span data-ttu-id="d4771-302">種類</span><span class="sxs-lookup"><span data-stu-id="d4771-302">Type</span></span>| <span data-ttu-id="d4771-303">説明</span><span class="sxs-lookup"><span data-stu-id="d4771-303">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="d4771-304">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="d4771-304">LocalClientTime</span></span>](/javascript/api/outlook_1_5/office.LocalClientTime)|<span data-ttu-id="d4771-305">変換するローカル時刻の値。</span><span class="sxs-lookup"><span data-stu-id="d4771-305">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d4771-306">要件</span><span class="sxs-lookup"><span data-stu-id="d4771-306">Requirements</span></span>

|<span data-ttu-id="d4771-307">要件</span><span class="sxs-lookup"><span data-stu-id="d4771-307">Requirement</span></span>| <span data-ttu-id="d4771-308">値</span><span class="sxs-lookup"><span data-stu-id="d4771-308">Value</span></span>|
|---|---|
|[<span data-ttu-id="d4771-309">最小限のメールボックス要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d4771-309">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d4771-310">1.0</span><span class="sxs-lookup"><span data-stu-id="d4771-310">1.0</span></span>|
|[<span data-ttu-id="d4771-311">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d4771-311">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d4771-312">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d4771-312">ReadItem</span></span>|
|[<span data-ttu-id="d4771-313">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d4771-313">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d4771-314">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="d4771-314">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d4771-315">戻り値 :</span><span class="sxs-lookup"><span data-stu-id="d4771-315">Returns:</span></span>

<span data-ttu-id="d4771-316">時間が UTC で表現された Date オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="d4771-316">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="d4771-317">

<dt>種類</dt>

</span><span class="sxs-lookup"><span data-stu-id="d4771-317">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="d4771-318">Date</span><span class="sxs-lookup"><span data-stu-id="d4771-318">Date</span></span></dd>

</dl>

####  <a name="displayappointmentformitemid"></a><span data-ttu-id="d4771-319">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="d4771-319">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="d4771-320">既存の予定表の予定を表示します。</span><span class="sxs-lookup"><span data-stu-id="d4771-320">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="d4771-321">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="d4771-321">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d4771-322">`displayAppointmentForm` メソッドは、デスクトップ上の新しいウィンドウまたはモバイル デバイス上のダイアログ ボックスで、既存の予定表の予定を開きます。</span><span class="sxs-lookup"><span data-stu-id="d4771-322">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="d4771-p112">Outlook for Mac では、このメソッドを使用して、定期的に繰り返される予定の一部ではない単発の予定、または定期的に繰り替えされる予定の元の予定を表示できます。ただし、一連の予定のインスタンスは表示できません。これは、Outlook for Mac では、定期的に繰り返されるインスタンスのプロパティ  (項目 ID を含む) にアクセスできないためです。</span><span class="sxs-lookup"><span data-stu-id="d4771-p112">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="d4771-325">Outlook Web App では、このメソッドで、指定されたフォームの本文が 32 KB 以下の文字数の場合に、そのフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="d4771-325">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="d4771-326">指定した項目識別子が既存の予定を識別しない場合には、クライアント コンピュータまたはデバイスで空白のウィンドウが開き、エラー メッセージは返されません。</span><span class="sxs-lookup"><span data-stu-id="d4771-326">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d4771-327">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="d4771-327">Parameters:</span></span>

|<span data-ttu-id="d4771-328">名前</span><span class="sxs-lookup"><span data-stu-id="d4771-328">Name</span></span>| <span data-ttu-id="d4771-329">種類</span><span class="sxs-lookup"><span data-stu-id="d4771-329">Type</span></span>| <span data-ttu-id="d4771-330">説明</span><span class="sxs-lookup"><span data-stu-id="d4771-330">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="d4771-331">文字列</span><span class="sxs-lookup"><span data-stu-id="d4771-331">String</span></span>|<span data-ttu-id="d4771-332">既存の予定表の予定の Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="d4771-332">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d4771-333">要件</span><span class="sxs-lookup"><span data-stu-id="d4771-333">Requirements</span></span>

|<span data-ttu-id="d4771-334">要件</span><span class="sxs-lookup"><span data-stu-id="d4771-334">Requirement</span></span>| <span data-ttu-id="d4771-335">値</span><span class="sxs-lookup"><span data-stu-id="d4771-335">Value</span></span>|
|---|---|
|[<span data-ttu-id="d4771-336">最小限のメールボックス要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d4771-336">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d4771-337">1.0</span><span class="sxs-lookup"><span data-stu-id="d4771-337">1.0</span></span>|
|[<span data-ttu-id="d4771-338">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d4771-338">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d4771-339">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d4771-339">ReadItem</span></span>|
|[<span data-ttu-id="d4771-340">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d4771-340">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d4771-341">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="d4771-341">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d4771-342">例</span><span class="sxs-lookup"><span data-stu-id="d4771-342">Example</span></span>

```
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  <a name="displaymessageformitemid"></a><span data-ttu-id="d4771-343">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="d4771-343">displayMessageForm(itemId)</span></span>

<span data-ttu-id="d4771-344">既存のメッセージを表示します。</span><span class="sxs-lookup"><span data-stu-id="d4771-344">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="d4771-345">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="d4771-345">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d4771-346">`displayMessageForm` メソッドは、デスクトップ上の新しいウィンドウまたはモバイル デバイス上のダイアログ ボックスで既存のメッセージを開きます。</span><span class="sxs-lookup"><span data-stu-id="d4771-346">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="d4771-347">Outlook Web App では、このメソッドは、指定されたフォームの本文が 32 KB 以下の文字数の場合に、そのフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="d4771-347">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="d4771-348">指定した項目識別子が既存のメッセージを識別しない場合には、クラアント コンピュータでメッセージは表示されず、エラー メッセージは返されません。</span><span class="sxs-lookup"><span data-stu-id="d4771-348">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="d4771-p113">予定を表す `displayMessageForm`  が含まれる `itemId`   を使用しないでください。`displayAppointmentForm` メソッドを使用して既存の予定を表示し、`displayNewAppointmentForm` を使用して新しい予定を作成するフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="d4771-p113">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d4771-351">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="d4771-351">Parameters:</span></span>

|<span data-ttu-id="d4771-352">名前</span><span class="sxs-lookup"><span data-stu-id="d4771-352">Name</span></span>| <span data-ttu-id="d4771-353">種類</span><span class="sxs-lookup"><span data-stu-id="d4771-353">Type</span></span>| <span data-ttu-id="d4771-354">説明</span><span class="sxs-lookup"><span data-stu-id="d4771-354">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="d4771-355">文字列</span><span class="sxs-lookup"><span data-stu-id="d4771-355">String</span></span>|<span data-ttu-id="d4771-356">既存のメッセージの Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="d4771-356">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d4771-357">要件</span><span class="sxs-lookup"><span data-stu-id="d4771-357">Requirements</span></span>

|<span data-ttu-id="d4771-358">要件</span><span class="sxs-lookup"><span data-stu-id="d4771-358">Requirement</span></span>| <span data-ttu-id="d4771-359">値</span><span class="sxs-lookup"><span data-stu-id="d4771-359">Value</span></span>|
|---|---|
|[<span data-ttu-id="d4771-360">最小限のメールボックス要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d4771-360">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d4771-361">1.0</span><span class="sxs-lookup"><span data-stu-id="d4771-361">1.0</span></span>|
|[<span data-ttu-id="d4771-362">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d4771-362">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d4771-363">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d4771-363">ReadItem</span></span>|
|[<span data-ttu-id="d4771-364">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d4771-364">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d4771-365">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="d4771-365">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d4771-366">例</span><span class="sxs-lookup"><span data-stu-id="d4771-366">Example</span></span>

```
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="d4771-367">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="d4771-367">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="d4771-368">新しい予定表の予定を作成するためのフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="d4771-368">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="d4771-369">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="d4771-369">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d4771-p114">`displayNewAppointmentForm` メソッドを使用すると、ユーザーが新しい予定または会議を作成できるフォームが開きます。パラメータを指定すると、予定のフォーム フィールドにパラメータの内容が自動的に入力されます。</span><span class="sxs-lookup"><span data-stu-id="d4771-p114">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="d4771-p115">Outlook Web App および OWA for Devices では、このメソッドは出席者フィールドが含まれるフォームを常に表示します。入力因数として出席者を指定しない場合には、このメソッドは [**保存**] ボタンのあるフォームを表示します。出席者を指定した場合には、フォームにはその出席者と [**送信**] ボタンが含まれます。</span><span class="sxs-lookup"><span data-stu-id="d4771-p115">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="d4771-p116">Outlook リッチ クライアントおよび Outlook RT では、`requiredAttendees`、`optionalAttendees` または `resources` パラメータに出席者またはリソースを指定した場合、このメソッドは [**送信**] ボタンがある会議フォームを表示します。受信者を指定しない場合には、このメソッドは [**保存して閉じる**] ボタンがある予定フォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="d4771-p116">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="d4771-377">パラメータのいずれかが指定されたサイズ制限を超えた場合、または不明なパラメータ名が指定された場合には、例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="d4771-377">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d4771-378">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="d4771-378">Parameters:</span></span>

|<span data-ttu-id="d4771-379">名前</span><span class="sxs-lookup"><span data-stu-id="d4771-379">Name</span></span>| <span data-ttu-id="d4771-380">種類</span><span class="sxs-lookup"><span data-stu-id="d4771-380">Type</span></span>| <span data-ttu-id="d4771-381">説明</span><span class="sxs-lookup"><span data-stu-id="d4771-381">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="d4771-382">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d4771-382">Object</span></span> | <span data-ttu-id="d4771-383">新しい予定を記述するパラメータの辞書。</span><span class="sxs-lookup"><span data-stu-id="d4771-383">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="d4771-384">配列。&lt;文字列&gt; | 配列です。&lt;[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="d4771-384">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="d4771-p117">予定への各必須出席者の電子メール アドレスを含む文字列の配列、または出席者の `EmailAddressDetails` オブジェクトを含む配列。この配列は、最大 100 エントリまでに制限されています。</span><span class="sxs-lookup"><span data-stu-id="d4771-p117">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="d4771-387">配列。&lt;文字列&gt; | 配列です。&lt;[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="d4771-387">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="d4771-p118">予定への各任意出席者の電子メール アドレスを含む文字列の配列、または出席者の `EmailAddressDetails` オブジェクトを含む配列。この配列は、最大 100 エントリまでに制限されています。</span><span class="sxs-lookup"><span data-stu-id="d4771-p118">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="d4771-390">Date</span><span class="sxs-lookup"><span data-stu-id="d4771-390">Date</span></span> | <span data-ttu-id="d4771-391">予定の開始日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="d4771-391">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="d4771-392">Date</span><span class="sxs-lookup"><span data-stu-id="d4771-392">Date</span></span> | <span data-ttu-id="d4771-393">予定の終了日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="d4771-393">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="d4771-394">文字列</span><span class="sxs-lookup"><span data-stu-id="d4771-394">String</span></span> | <span data-ttu-id="d4771-p119">予定の場所を含む文字列。文字列は最大 255 文字までに制限されています。</span><span class="sxs-lookup"><span data-stu-id="d4771-p119">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="d4771-397">配列。&lt; 文字列&gt;</span><span class="sxs-lookup"><span data-stu-id="d4771-397">Array.&lt;String&gt;</span></span> | <span data-ttu-id="d4771-p120">予定に必要なリソースを含む文字列の配列。配列は最大 100 エントリまでに制限されています。</span><span class="sxs-lookup"><span data-stu-id="d4771-p120">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="d4771-400">文字列</span><span class="sxs-lookup"><span data-stu-id="d4771-400">String</span></span> | <span data-ttu-id="d4771-p121">予定の件名を含む文字列。文字列は最大 255 文字までに制限されています。</span><span class="sxs-lookup"><span data-stu-id="d4771-p121">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="d4771-403">文字列</span><span class="sxs-lookup"><span data-stu-id="d4771-403">String</span></span> | <span data-ttu-id="d4771-p122">予定の本文。本文の内容は、最大サイズが 32 KB までに制限されています。</span><span class="sxs-lookup"><span data-stu-id="d4771-p122">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="d4771-406">要件</span><span class="sxs-lookup"><span data-stu-id="d4771-406">Requirements</span></span>

|<span data-ttu-id="d4771-407">要件</span><span class="sxs-lookup"><span data-stu-id="d4771-407">Requirement</span></span>| <span data-ttu-id="d4771-408">値</span><span class="sxs-lookup"><span data-stu-id="d4771-408">Value</span></span>|
|---|---|
|[<span data-ttu-id="d4771-409">最小限のメールボックス要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d4771-409">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d4771-410">1.0</span><span class="sxs-lookup"><span data-stu-id="d4771-410">1.0</span></span>|
|[<span data-ttu-id="d4771-411">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d4771-411">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d4771-412">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d4771-412">ReadItem</span></span>|
|[<span data-ttu-id="d4771-413">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d4771-413">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d4771-414">読み取り</span><span class="sxs-lookup"><span data-stu-id="d4771-414">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d4771-415">例</span><span class="sxs-lookup"><span data-stu-id="d4771-415">Example</span></span>

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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="d4771-416">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="d4771-416">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="d4771-417">REST API または Exchange Web サービスを呼び出すために使用するトークンを含む文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="d4771-417">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="d4771-p123">`getCallbackTokenAsync` メソッドは、非同期の呼び出しを行なって、ユーザーのメールボックスをホストする Exchange Server から opaque トークンを取得します。コールバック トークンの有効期間は 5 分です。</span><span class="sxs-lookup"><span data-stu-id="d4771-p123">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="d4771-420">可能な場合は常に、アドインでは Exchange Web サービスの代わりに REST API を使用することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="d4771-420">Note: It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span> 

<span data-ttu-id="d4771-421">**REST トークン**</span><span class="sxs-lookup"><span data-stu-id="d4771-421">**REST Tokens**</span></span>

<span data-ttu-id="d4771-p124">REST トークンが要求された場合 (`options.isRest = true`) には、作成されたトークンは Exchange Web サービスの呼び出しを認証するためには機能しません。このトークンは、アドインがマニフェストで [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) アクセス許可を指定しない限り、現在の項目およびその添付ファイルへの読み取り専用の範囲に制限されます。`ReadWriteMailbox` アクセス許可が指定された場合には、作成されるトークンは、メールを送信する機能など、メール、予定表、連絡先への読み取り/書き込みアクセスを付与します。</span><span class="sxs-lookup"><span data-stu-id="d4771-p124">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="d4771-425">アドインでは、`restUrl`プロパティを使用して、REST API 呼び出しを行うときに使用する正しい URL を決定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="d4771-425">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="d4771-426">**EWS トークン**</span><span class="sxs-lookup"><span data-stu-id="d4771-426">**EWS Tokens**</span></span>

<span data-ttu-id="d4771-p125">EWS トークンが要求された場合(`options.isRest = false`) には、作成されるトークンは REST API の呼び出しを認証するためには機能しません。このトークンは、現在の項目にアクセスできる範囲に制限されます。</span><span class="sxs-lookup"><span data-stu-id="d4771-p125">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="d4771-429">アドインでは、`ewsUrl` プロパティを使用して、EWS 呼び出しを行うときに使用する正しい URL を決定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="d4771-429">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d4771-430">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="d4771-430">Parameters:</span></span>

|<span data-ttu-id="d4771-431">名前</span><span class="sxs-lookup"><span data-stu-id="d4771-431">Name</span></span>| <span data-ttu-id="d4771-432">種類</span><span class="sxs-lookup"><span data-stu-id="d4771-432">Type</span></span>| <span data-ttu-id="d4771-433">属性</span><span class="sxs-lookup"><span data-stu-id="d4771-433">Attributes</span></span>| <span data-ttu-id="d4771-434">説明</span><span class="sxs-lookup"><span data-stu-id="d4771-434">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="d4771-435">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d4771-435">Object</span></span> | <span data-ttu-id="d4771-436">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="d4771-436">&lt;optional&gt;</span></span> | <span data-ttu-id="d4771-437">以下のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="d4771-437">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="d4771-438">ブール値</span><span class="sxs-lookup"><span data-stu-id="d4771-438">Boolean</span></span> |  <span data-ttu-id="d4771-439">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="d4771-439">&lt;optional&gt;</span></span> | <span data-ttu-id="d4771-p126">提供されたトークンを Outlook REST API または Exchange Web サービスに使用するかどうかを決定します。既定値は、`false`です。</span><span class="sxs-lookup"><span data-stu-id="d4771-p126">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="d4771-442">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d4771-442">Object</span></span> |  <span data-ttu-id="d4771-443">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="d4771-443">&lt;optional&gt;</span></span> | <span data-ttu-id="d4771-444">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="d4771-444">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="d4771-445">関数</span><span class="sxs-lookup"><span data-stu-id="d4771-445">function</span></span>||<span data-ttu-id="d4771-p127">メソッドが完了すると、`callback` パラメータに渡された関数が、単一パラメータ `asyncResult` で呼び出されます。このパラメータは、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトです。トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="d4771-p127">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d4771-448">要件</span><span class="sxs-lookup"><span data-stu-id="d4771-448">Requirements</span></span>

|<span data-ttu-id="d4771-449">要件</span><span class="sxs-lookup"><span data-stu-id="d4771-449">Requirement</span></span>| <span data-ttu-id="d4771-450">値</span><span class="sxs-lookup"><span data-stu-id="d4771-450">Value</span></span>|
|---|---|
|[<span data-ttu-id="d4771-451">最小限のメールボックス要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d4771-451">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d4771-452">1.5</span><span class="sxs-lookup"><span data-stu-id="d4771-452">1.5</span></span> |
|[<span data-ttu-id="d4771-453">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d4771-453">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d4771-454">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d4771-454">ReadItem</span></span>|
|[<span data-ttu-id="d4771-455">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d4771-455">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d4771-456">新規作成と閲覧</span><span class="sxs-lookup"><span data-stu-id="d4771-456">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="d4771-457">例</span><span class="sxs-lookup"><span data-stu-id="d4771-457">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="d4771-458">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="d4771-458">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="d4771-459">Exchange Server から添付ファイルやアイテムを取得するために使用するトークンを含む文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="d4771-459">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="d4771-p128">`getCallbackTokenAsync` メソッドは、非同期の呼び出しを行なって、ユーザーのメールボックスをホストする Exchange Server から opaque トークンを取得します。コールバック トークンの有効期間は 5 分です。</span><span class="sxs-lookup"><span data-stu-id="d4771-p128">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="d4771-p129">このトークンと、添付ファイル識別子または項目識別子は、サードパーティーのシステムに渡すことができます。サードパーティーのシステムでは、添付ファイルまたは項目を返すための Exchange Web サービス (EWS) の [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) または [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) 操作を呼び出すのに、トークンをベアラー承認トークンとして使用します。たとえば、リモート サービスを作成して[選択した項目から添付ファイルを取得](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item)することができます。</span><span class="sxs-lookup"><span data-stu-id="d4771-p129">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="d4771-465">アプリでは、閲覧モードで \*\*  メソッドを呼び出すために、 \*\*  ReadItem`getCallbackTokenAsync`  アクセス許可をアプリのマニフェストで指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="d4771-465">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="d4771-p130">新規作成モードでは、[ `saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出して、`getCallbackTokenAsync` メソッドに渡すための項目識別子を取得する必要があります。アプリには、\*\*  メソッドを呼び出すために \*\* ReadWriteItem`saveAsync`   アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="d4771-p130">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d4771-468">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="d4771-468">Parameters:</span></span>

|<span data-ttu-id="d4771-469">名前</span><span class="sxs-lookup"><span data-stu-id="d4771-469">Name</span></span>| <span data-ttu-id="d4771-470">種類</span><span class="sxs-lookup"><span data-stu-id="d4771-470">Type</span></span>| <span data-ttu-id="d4771-471">属性</span><span class="sxs-lookup"><span data-stu-id="d4771-471">Attributes</span></span>| <span data-ttu-id="d4771-472">説明</span><span class="sxs-lookup"><span data-stu-id="d4771-472">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="d4771-473">関数</span><span class="sxs-lookup"><span data-stu-id="d4771-473">function</span></span>||<span data-ttu-id="d4771-p131">メソッドが完了すると、`callback` パラメータに渡された関数が、単一パラメータ `asyncResult` で呼び出されます。このパラメータは、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトです。トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="d4771-p131">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="d4771-476">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d4771-476">Object</span></span>| <span data-ttu-id="d4771-477">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="d4771-477">&lt;optional&gt;</span></span>|<span data-ttu-id="d4771-478">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="d4771-478">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d4771-479">要件</span><span class="sxs-lookup"><span data-stu-id="d4771-479">Requirements</span></span>

|<span data-ttu-id="d4771-480">要件</span><span class="sxs-lookup"><span data-stu-id="d4771-480">Requirement</span></span>| <span data-ttu-id="d4771-481">値</span><span class="sxs-lookup"><span data-stu-id="d4771-481">Value</span></span>|
|---|---|
|[<span data-ttu-id="d4771-482">最小限のメールボックス要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d4771-482">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d4771-483">1.3</span><span class="sxs-lookup"><span data-stu-id="d4771-483">1.3</span></span>|
|[<span data-ttu-id="d4771-484">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d4771-484">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d4771-485">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d4771-485">ReadItem</span></span>|
|[<span data-ttu-id="d4771-486">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d4771-486">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d4771-487">新規作成と閲覧</span><span class="sxs-lookup"><span data-stu-id="d4771-487">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="d4771-488">例</span><span class="sxs-lookup"><span data-stu-id="d4771-488">Example</span></span>

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="d4771-489">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="d4771-489">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="d4771-490">ユーザーと Office アドインを識別するトークンを取得します。</span><span class="sxs-lookup"><span data-stu-id="d4771-490">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="d4771-491">`getUserIdentityTokenAsync` メソッドは、[アドインとユーザーをサードパーティのシステムで識別して認証](https://docs.microsoft.com/outlook/add-ins/authentication)するのに使用できるトークンを返します。</span><span class="sxs-lookup"><span data-stu-id="d4771-491">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](https://docs.microsoft.com/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="d4771-492">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="d4771-492">Parameters:</span></span>

|<span data-ttu-id="d4771-493">名前</span><span class="sxs-lookup"><span data-stu-id="d4771-493">Name</span></span>| <span data-ttu-id="d4771-494">種類</span><span class="sxs-lookup"><span data-stu-id="d4771-494">Type</span></span>| <span data-ttu-id="d4771-495">属性</span><span class="sxs-lookup"><span data-stu-id="d4771-495">Attributes</span></span>| <span data-ttu-id="d4771-496">説明</span><span class="sxs-lookup"><span data-stu-id="d4771-496">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="d4771-497">関数</span><span class="sxs-lookup"><span data-stu-id="d4771-497">function</span></span>||<span data-ttu-id="d4771-498">メソッドが完了すると、`callback` パラメータで渡された関数が、単一パラメータ `asyncResult` で呼び出されます。このパラメータは、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="d4771-498">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="d4771-499">トークンは、`asyncResult.value`プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="d4771-499">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="d4771-500">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d4771-500">Object</span></span>| <span data-ttu-id="d4771-501">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="d4771-501">&lt;optional&gt;</span></span>|<span data-ttu-id="d4771-502">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="d4771-502">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d4771-503">要件</span><span class="sxs-lookup"><span data-stu-id="d4771-503">Requirements</span></span>

|<span data-ttu-id="d4771-504">要件</span><span class="sxs-lookup"><span data-stu-id="d4771-504">Requirement</span></span>| <span data-ttu-id="d4771-505">値</span><span class="sxs-lookup"><span data-stu-id="d4771-505">Value</span></span>|
|---|---|
|[<span data-ttu-id="d4771-506">最小限のメールボックス要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d4771-506">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d4771-507">1.0</span><span class="sxs-lookup"><span data-stu-id="d4771-507">1.0</span></span>|
|[<span data-ttu-id="d4771-508">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d4771-508">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d4771-509">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d4771-509">ReadItem</span></span>|
|[<span data-ttu-id="d4771-510">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d4771-510">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d4771-511">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="d4771-511">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d4771-512">例</span><span class="sxs-lookup"><span data-stu-id="d4771-512">Example</span></span>

```js
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="d4771-513">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="d4771-513">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="d4771-514">ユーザーのメールボックスをホストしている Exchange Server上の Exchange Web サービス (EWS) のサービスに対して非同期の要求を行います。</span><span class="sxs-lookup"><span data-stu-id="d4771-514">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="d4771-515">このメソッドは、次のシナリオではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="d4771-515">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="d4771-516">Outlook for iOS または Outlook for Android で</span><span class="sxs-lookup"><span data-stu-id="d4771-516">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="d4771-517">アドインが Gmail のメールボックスにロードされる場合</span><span class="sxs-lookup"><span data-stu-id="d4771-517">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="d4771-518">これらの場合、アドインではユーザーのメールボックスにアクセスするために、代わりに [REST API を使用する](https://docs.microsoft.com/outlook/add-ins/use-rest-api)必要があります。</span><span class="sxs-lookup"><span data-stu-id="d4771-518">In these cases, add-ins should [use REST APIs](https://docs.microsoft.com/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="d4771-519">`makeEwsRequestAsync` メソッドは、アドインの代わりに Exchange に EWS 要求を送信します。</span><span class="sxs-lookup"><span data-stu-id="d4771-519">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="d4771-520">サポートされている EWS 操作の一覧については、 「[ Outlook アドインから Web サービスを呼び出す](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="d4771-520">See [Call web services from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="d4771-521">`makeEwsRequestAsync` メソッドで、フォルダー関連アイテムを要求することはできません。</span><span class="sxs-lookup"><span data-stu-id="d4771-521">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="d4771-522">XML 要求では、UTF-8 エンコードを指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="d4771-522">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="d4771-p133">アドインには、\*\*  メソッドを使用するために \*\* ReadWriteMailbox`makeEwsRequestAsync`  アクセス許可が必要です。**ReadWriteMailbox** アクセス許可と、`makeEwsRequestAsync` メソッドで呼び出すことのできる EWS 操作の使用の詳細については、「[ユーザーのメールボックスへのメール アドイン アクセスのアクセス許可を指定する](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="d4771-p133">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="d4771-525">サーバー管理者は、クライアント アクセス サーバー の EWS ディレクトリ上で `OAuthAuthentication` を true に設定して、`makeEwsRequestAsync` メソッドで EWS 要求を行えるようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="d4771-525">NOTE: The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="d4771-526">バージョンの相違点</span><span class="sxs-lookup"><span data-stu-id="d4771-526">Version differences</span></span>

<span data-ttu-id="d4771-527">バージョン 15.0.4535.1004 より前のバージョンの Outlook で実行しているメール アプリで `makeEwsRequestAsync` メソッドを使用する場合には、エンコード値を `ISO-8859-1` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="d4771-527">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="d4771-p134">メール アプリが Outlook on the web で実行されている場合には、エンコード値を設定する必要はありません。メールボックスを使用してメール アプリが Outlook で実行されているのか、Outlook on the web で実行されているのかを判断する必要があります。mailbox.diagnostics.hostVersion プロパティを使用すれば、どのバージョンの Outlook が実行されているのかがわかります。</span><span class="sxs-lookup"><span data-stu-id="d4771-p134">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d4771-531">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="d4771-531">Parameters:</span></span>

|<span data-ttu-id="d4771-532">名前</span><span class="sxs-lookup"><span data-stu-id="d4771-532">Name</span></span>| <span data-ttu-id="d4771-533">種類</span><span class="sxs-lookup"><span data-stu-id="d4771-533">Type</span></span>| <span data-ttu-id="d4771-534">属性</span><span class="sxs-lookup"><span data-stu-id="d4771-534">Attributes</span></span>| <span data-ttu-id="d4771-535">説明</span><span class="sxs-lookup"><span data-stu-id="d4771-535">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="d4771-536">文字列</span><span class="sxs-lookup"><span data-stu-id="d4771-536">String</span></span>||<span data-ttu-id="d4771-537">EWS 要求。</span><span class="sxs-lookup"><span data-stu-id="d4771-537">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="d4771-538">関数</span><span class="sxs-lookup"><span data-stu-id="d4771-538">function</span></span>||<span data-ttu-id="d4771-539">メソッドが完了すると、`callback` パラメータで渡された関数が、単一パラメータ `asyncResult` で呼び出されます。このパラメータは、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="d4771-539">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="d4771-540">EWS 呼び出しの XML 結果は、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="d4771-540">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="d4771-541">結果のサイズが 1 MB を超えている場合、エラー メッセージが返されます。</span><span class="sxs-lookup"><span data-stu-id="d4771-541">If the result exceeds 1 MB in size, an error message is returned instead.| | Object| optional|Any state data that is passed to the asynchronous method.|</span></span>|
|`userContext`| <span data-ttu-id="d4771-542">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d4771-542">Object</span></span>| <span data-ttu-id="d4771-543">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="d4771-543">&lt;optional&gt;</span></span>|<span data-ttu-id="d4771-544">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="d4771-544">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d4771-545">要件</span><span class="sxs-lookup"><span data-stu-id="d4771-545">Requirements</span></span>

|<span data-ttu-id="d4771-546">要件</span><span class="sxs-lookup"><span data-stu-id="d4771-546">Requirement</span></span>| <span data-ttu-id="d4771-547">値</span><span class="sxs-lookup"><span data-stu-id="d4771-547">Value</span></span>|
|---|---|
|[<span data-ttu-id="d4771-548">最小限のメールボックス要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d4771-548">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d4771-549">1.0</span><span class="sxs-lookup"><span data-stu-id="d4771-549">1.0</span></span>|
|[<span data-ttu-id="d4771-550">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d4771-550">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d4771-551">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="d4771-551">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="d4771-552">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d4771-552">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d4771-553">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="d4771-553">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d4771-554">例</span><span class="sxs-lookup"><span data-stu-id="d4771-554">Example</span></span>

<span data-ttu-id="d4771-555">次の例は、`makeEwsRequestAsync` を呼び出し、`GetItem` 操作を使って項目の件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="d4771-555">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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