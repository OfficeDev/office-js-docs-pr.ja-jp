---
title: Office.context.mailbox - 要件セット 1.6
description: ''
ms.date: 10/31/2018
ms.openlocfilehash: cf7c5620d9109f2350972e0f797e7f195f91a90e
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433349"
---
# <a name="mailbox"></a><span data-ttu-id="f5953-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="f5953-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="f5953-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="f5953-103">Office.context.mailbox</span></span>

<span data-ttu-id="f5953-104">Microsoft Outlook と Microsoft Outlook on the web の Outlook アドイン オブジェクト モデルへのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="f5953-104">Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f5953-105">要件</span><span class="sxs-lookup"><span data-stu-id="f5953-105">Requirements</span></span>

|<span data-ttu-id="f5953-106">要件</span><span class="sxs-lookup"><span data-stu-id="f5953-106">Requirement</span></span>| <span data-ttu-id="f5953-107">値</span><span class="sxs-lookup"><span data-stu-id="f5953-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="f5953-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f5953-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f5953-109">1.0</span><span class="sxs-lookup"><span data-stu-id="f5953-109">1.0</span></span>|
|[<span data-ttu-id="f5953-110">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f5953-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f5953-111">制限あり</span><span class="sxs-lookup"><span data-stu-id="f5953-111">Restricted</span></span>|
|[<span data-ttu-id="f5953-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f5953-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f5953-113">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="f5953-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="f5953-114">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="f5953-114">Members and methods</span></span>

| <span data-ttu-id="f5953-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="f5953-115">Member</span></span> | <span data-ttu-id="f5953-116">種類</span><span class="sxs-lookup"><span data-stu-id="f5953-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="f5953-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="f5953-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="f5953-118">メンバー</span><span class="sxs-lookup"><span data-stu-id="f5953-118">Member</span></span> |
| [<span data-ttu-id="f5953-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="f5953-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="f5953-120">メンバー</span><span class="sxs-lookup"><span data-stu-id="f5953-120">Member</span></span> |
| [<span data-ttu-id="f5953-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="f5953-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="f5953-122">メソッド</span><span class="sxs-lookup"><span data-stu-id="f5953-122">Method</span></span> |
| [<span data-ttu-id="f5953-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="f5953-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="f5953-124">メソッド</span><span class="sxs-lookup"><span data-stu-id="f5953-124">Method</span></span> |
| [<span data-ttu-id="f5953-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="f5953-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) | <span data-ttu-id="f5953-126">メソッド</span><span class="sxs-lookup"><span data-stu-id="f5953-126">Method</span></span> |
| [<span data-ttu-id="f5953-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="f5953-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="f5953-128">メソッド</span><span class="sxs-lookup"><span data-stu-id="f5953-128">Method</span></span> |
| [<span data-ttu-id="f5953-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="f5953-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="f5953-130">メソッド</span><span class="sxs-lookup"><span data-stu-id="f5953-130">Method</span></span> |
| [<span data-ttu-id="f5953-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="f5953-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="f5953-132">メソッド</span><span class="sxs-lookup"><span data-stu-id="f5953-132">Method</span></span> |
| [<span data-ttu-id="f5953-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="f5953-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="f5953-134">メソッド</span><span class="sxs-lookup"><span data-stu-id="f5953-134">Method</span></span> |
| [<span data-ttu-id="f5953-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="f5953-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="f5953-136">メソッド</span><span class="sxs-lookup"><span data-stu-id="f5953-136">Method</span></span> |
| [<span data-ttu-id="f5953-137">displayNewMessageForm</span><span class="sxs-lookup"><span data-stu-id="f5953-137">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="f5953-138">メソッド</span><span class="sxs-lookup"><span data-stu-id="f5953-138">Method</span></span> |
| [<span data-ttu-id="f5953-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="f5953-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="f5953-140">メソッド</span><span class="sxs-lookup"><span data-stu-id="f5953-140">Method</span></span> |
| [<span data-ttu-id="f5953-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="f5953-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="f5953-142">メソッド</span><span class="sxs-lookup"><span data-stu-id="f5953-142">Method</span></span> |
| [<span data-ttu-id="f5953-143">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="f5953-143">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="f5953-144">メソッド</span><span class="sxs-lookup"><span data-stu-id="f5953-144">Method</span></span> |
| [<span data-ttu-id="f5953-145">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="f5953-145">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="f5953-146">メソッド</span><span class="sxs-lookup"><span data-stu-id="f5953-146">Method</span></span> |
| [<span data-ttu-id="f5953-147">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="f5953-147">removeHandlerAsync</span></span>](#removehandlerasynceventtype-handler-options-callback) | <span data-ttu-id="f5953-148">メソッド</span><span class="sxs-lookup"><span data-stu-id="f5953-148">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="f5953-149">名前空間</span><span class="sxs-lookup"><span data-stu-id="f5953-149">Namespaces</span></span>

<span data-ttu-id="f5953-150">[diagnostics](Office.context.mailbox.diagnostics.md):Outlook アドインに診断情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="f5953-150">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="f5953-151">[item](Office.context.mailbox.item.md):Outlook アドインのメッセージや予定にアクセスするためのメソッドとプロパティを提供します。</span><span class="sxs-lookup"><span data-stu-id="f5953-151">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="f5953-152">[userProfile](Office.context.mailbox.userProfile.md):Outlook アドインのユーザーに関する情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="f5953-152">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="f5953-153">メンバー</span><span class="sxs-lookup"><span data-stu-id="f5953-153">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="f5953-154">ewsUrl :String</span><span class="sxs-lookup"><span data-stu-id="f5953-154">ewsUrl :String</span></span>

<span data-ttu-id="f5953-p101">このメール アカウントの Exchange Web サービス (EWS) エンドポイントの URL を取得します。読み取りモードのみです。</span><span class="sxs-lookup"><span data-stu-id="f5953-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="f5953-157">このメンバーは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f5953-157">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="f5953-p102">
  `ewsUrl\` 値は、リモート サービスで、ユーザーのメールボックスに EWS 呼び出しを行うために使うことができます。たとえば、[選択したアイテムから添付ファイルを取得する](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="f5953-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="f5953-160">アプリが閲覧モードで `ewsUrl` メンバーを呼び出すには、アプリのマニフェスト内に **ReadItem** アクセス許可が指定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="f5953-160">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="f5953-p103">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出してから、`ewsUrl` メンバーを使用する必要があります。アプリには、`saveAsync` メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="f5953-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="f5953-163">型:</span><span class="sxs-lookup"><span data-stu-id="f5953-163">Type:</span></span>

*   <span data-ttu-id="f5953-164">String</span><span class="sxs-lookup"><span data-stu-id="f5953-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f5953-165">要件</span><span class="sxs-lookup"><span data-stu-id="f5953-165">Requirements</span></span>

|<span data-ttu-id="f5953-166">要件</span><span class="sxs-lookup"><span data-stu-id="f5953-166">Requirement</span></span>| <span data-ttu-id="f5953-167">値</span><span class="sxs-lookup"><span data-stu-id="f5953-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="f5953-168">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f5953-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f5953-169">1.0</span><span class="sxs-lookup"><span data-stu-id="f5953-169">1.0</span></span>|
|[<span data-ttu-id="f5953-170">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f5953-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f5953-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f5953-171">ReadItem</span></span>|
|[<span data-ttu-id="f5953-172">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f5953-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f5953-173">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="f5953-173">Compose or read</span></span>|

#### <a name="resturl-string"></a><span data-ttu-id="f5953-174">restUrl :String</span><span class="sxs-lookup"><span data-stu-id="f5953-174">restUrl :String</span></span>

<span data-ttu-id="f5953-175">この電子メール アカウントの REST エンドポイントの URL を取得します。</span><span class="sxs-lookup"><span data-stu-id="f5953-175">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="f5953-176">`restUrl` 値は、ユーザーのメールボックスに [REST API](https://docs.microsoft.com/outlook/rest/) 呼び出しを行うために使用できます。</span><span class="sxs-lookup"><span data-stu-id="f5953-176">The `restUrl` value can be used to make [REST API](https://docs.microsoft.com/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="f5953-177">アプリが閲覧モードで `restUrl` メンバーを呼び出すには、アプリのマニフェスト内に **ReadItem** アクセス許可が指定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="f5953-177">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="f5953-p104">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出してから、`restUrl` メンバーを使用する必要があります。アプリには、`saveAsync` メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="f5953-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="f5953-180">型:</span><span class="sxs-lookup"><span data-stu-id="f5953-180">Type:</span></span>

*   <span data-ttu-id="f5953-181">String</span><span class="sxs-lookup"><span data-stu-id="f5953-181">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f5953-182">要件</span><span class="sxs-lookup"><span data-stu-id="f5953-182">Requirements</span></span>

|<span data-ttu-id="f5953-183">要件</span><span class="sxs-lookup"><span data-stu-id="f5953-183">Requirement</span></span>| <span data-ttu-id="f5953-184">値</span><span class="sxs-lookup"><span data-stu-id="f5953-184">Value</span></span>|
|---|---|
|[<span data-ttu-id="f5953-185">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f5953-185">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f5953-186">1.5</span><span class="sxs-lookup"><span data-stu-id="f5953-186">1.5</span></span> |
|[<span data-ttu-id="f5953-187">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f5953-187">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f5953-188">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f5953-188">ReadItem</span></span>|
|[<span data-ttu-id="f5953-189">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f5953-189">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f5953-190">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="f5953-190">Compose or read</span></span>|

### <a name="methods"></a><span data-ttu-id="f5953-191">メソッド</span><span class="sxs-lookup"><span data-stu-id="f5953-191">Methods</span></span>

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="f5953-192">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="f5953-192">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="f5953-193">サポートされているイベントのイベント ハンドラーを追加します。</span><span class="sxs-lookup"><span data-stu-id="f5953-193">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="f5953-194">現在サポートされている唯一のイベントの種類は `Office.EventType.ItemChanged` です。これはユーザーが新しいアイテムを選択したときに呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="f5953-194">Currently the only supported event type is `Office.EventType.ItemChanged`, which is invoked when the user selects a new item.</span></span> <span data-ttu-id="f5953-195">このイベントは、ピン留め可能な作業ウィンドウを実装するアドインで使用され、現在選択されているアイテムに基づいて作業ウィンドウ UI をアドインで更新できるようにします。</span><span class="sxs-lookup"><span data-stu-id="f5953-195">This event is used by add-ins that implement a pinnable task pane, and allows the add-in to refresh the task pane UI based on the currently selected item.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f5953-196">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="f5953-196">Parameters:</span></span>

| <span data-ttu-id="f5953-197">名前</span><span class="sxs-lookup"><span data-stu-id="f5953-197">Name</span></span> | <span data-ttu-id="f5953-198">型</span><span class="sxs-lookup"><span data-stu-id="f5953-198">Type</span></span> | <span data-ttu-id="f5953-199">属性</span><span class="sxs-lookup"><span data-stu-id="f5953-199">Attributes</span></span> | <span data-ttu-id="f5953-200">説明</span><span class="sxs-lookup"><span data-stu-id="f5953-200">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="f5953-201">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="f5953-201">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="f5953-202">ハンドラーを呼び出す必要のあるイベント。</span><span class="sxs-lookup"><span data-stu-id="f5953-202">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="f5953-203">Function</span><span class="sxs-lookup"><span data-stu-id="f5953-203">Function</span></span> || <span data-ttu-id="f5953-p106">イベントを処理する関数。関数は、オブジェクト リテラルである単一パラメーターを受け入れる必要があります。パラメーターの `type` プロパティは、`addHandlerAsync` に渡される `eventType` パラメーターと一致します。</span><span class="sxs-lookup"><span data-stu-id="f5953-p106">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="f5953-207">Object</span><span class="sxs-lookup"><span data-stu-id="f5953-207">Object</span></span> | <span data-ttu-id="f5953-208">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f5953-208">&lt;optional&gt;</span></span> | <span data-ttu-id="f5953-209">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="f5953-209">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="f5953-210">Object</span><span class="sxs-lookup"><span data-stu-id="f5953-210">Object</span></span> | <span data-ttu-id="f5953-211">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f5953-211">&lt;optional&gt;</span></span> | <span data-ttu-id="f5953-212">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="f5953-212">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="f5953-213">function</span><span class="sxs-lookup"><span data-stu-id="f5953-213">function</span></span>| <span data-ttu-id="f5953-214">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f5953-214">&lt;optional&gt;</span></span>|<span data-ttu-id="f5953-215">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="f5953-215">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f5953-216">要件</span><span class="sxs-lookup"><span data-stu-id="f5953-216">Requirements</span></span>

|<span data-ttu-id="f5953-217">要件</span><span class="sxs-lookup"><span data-stu-id="f5953-217">Requirement</span></span>| <span data-ttu-id="f5953-218">値</span><span class="sxs-lookup"><span data-stu-id="f5953-218">Value</span></span>|
|---|---|
|[<span data-ttu-id="f5953-219">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f5953-219">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f5953-220">1.5</span><span class="sxs-lookup"><span data-stu-id="f5953-220">1.5</span></span> |
|[<span data-ttu-id="f5953-221">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f5953-221">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f5953-222">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f5953-222">ReadItem</span></span> |
|[<span data-ttu-id="f5953-223">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f5953-223">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f5953-224">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="f5953-224">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f5953-225">例</span><span class="sxs-lookup"><span data-stu-id="f5953-225">Example</span></span>

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

####  <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="f5953-226">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="f5953-226">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="f5953-227">REST 形式のアイテム ID を EWS 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="f5953-227">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="f5953-228">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f5953-228">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="f5953-p107">REST API ([Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) や [Microsoft Graph](https://graph.microsoft.io/) など) で取得されたアイテム ID は、Exchange Web サービス (EWS) に使用される形式とは異なる形式を使用します。`convertToEwsId` メソッドは、REST 形式の ID を EWS 用の適切な形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="f5953-p107">Item IDs retrieved via a REST API (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f5953-231">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="f5953-231">Parameters:</span></span>

|<span data-ttu-id="f5953-232">名前</span><span class="sxs-lookup"><span data-stu-id="f5953-232">Name</span></span>| <span data-ttu-id="f5953-233">型</span><span class="sxs-lookup"><span data-stu-id="f5953-233">Type</span></span>| <span data-ttu-id="f5953-234">説明</span><span class="sxs-lookup"><span data-stu-id="f5953-234">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="f5953-235">String</span><span class="sxs-lookup"><span data-stu-id="f5953-235">String</span></span>|<span data-ttu-id="f5953-236">Outlook REST API 形式のアイテム ID</span><span class="sxs-lookup"><span data-stu-id="f5953-236">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="f5953-237">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="f5953-237">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.restversion)|<span data-ttu-id="f5953-238">アイテム ID の取得に使用された Outlook REST API のバージョンを示す値。</span><span class="sxs-lookup"><span data-stu-id="f5953-238">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f5953-239">要件</span><span class="sxs-lookup"><span data-stu-id="f5953-239">Requirements</span></span>

|<span data-ttu-id="f5953-240">要件</span><span class="sxs-lookup"><span data-stu-id="f5953-240">Requirement</span></span>| <span data-ttu-id="f5953-241">値</span><span class="sxs-lookup"><span data-stu-id="f5953-241">Value</span></span>|
|---|---|
|[<span data-ttu-id="f5953-242">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f5953-242">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f5953-243">1.3</span><span class="sxs-lookup"><span data-stu-id="f5953-243">1.3</span></span>|
|[<span data-ttu-id="f5953-244">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f5953-244">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f5953-245">制限あり</span><span class="sxs-lookup"><span data-stu-id="f5953-245">Restricted</span></span>|
|[<span data-ttu-id="f5953-246">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f5953-246">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f5953-247">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="f5953-247">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f5953-248">戻り値:</span><span class="sxs-lookup"><span data-stu-id="f5953-248">Returns:</span></span>

<span data-ttu-id="f5953-249">型:String</span><span class="sxs-lookup"><span data-stu-id="f5953-249">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="f5953-250">例</span><span class="sxs-lookup"><span data-stu-id="f5953-250">Example</span></span>

```js
// Get an item's ID from a REST API
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the
// Outlook Mail API
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime"></a><span data-ttu-id="f5953-251">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="f5953-251">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)}</span></span>

<span data-ttu-id="f5953-252">クライアントのローカル時間で時間情報が含まれている辞書を取得します。</span><span class="sxs-lookup"><span data-stu-id="f5953-252">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="f5953-p108">Outlook 用メール アプリや Outlook Web App で使う日付と時刻では、異なるタイム ゾーンを使うことができます。Outlook では、クライアント コンピューターのタイム ゾーンを使います。Outlook Web App では、Exchange 管理センター (EAC) で設定されたタイム ゾーンを使います。ユーザー インターフェイスに表示される値が、常にユーザーが期待するタイム ゾーンと一致するように日付と時刻の値を処理する必要があります。</span><span class="sxs-lookup"><span data-stu-id="f5953-p108">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="f5953-p109">Outlook でメール アプリが実行されている場合、`convertToLocalClientTime` メソッドは、クライアント コンピューターのタイム ゾーンに設定された値のディクショナリ オブジェクトを返します。Outlook Web Apps でメール アプリが実行されている場合、`convertToLocalClientTime` メソッドは、EAC に指定されたタイム ゾーンに設定された値のディクショナリ オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="f5953-p109">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f5953-258">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="f5953-258">Parameters:</span></span>

|<span data-ttu-id="f5953-259">名前</span><span class="sxs-lookup"><span data-stu-id="f5953-259">Name</span></span>| <span data-ttu-id="f5953-260">型</span><span class="sxs-lookup"><span data-stu-id="f5953-260">Type</span></span>| <span data-ttu-id="f5953-261">説明</span><span class="sxs-lookup"><span data-stu-id="f5953-261">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="f5953-262">Date</span><span class="sxs-lookup"><span data-stu-id="f5953-262">Date</span></span>|<span data-ttu-id="f5953-263">日付オブジェクト</span><span class="sxs-lookup"><span data-stu-id="f5953-263">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f5953-264">要件</span><span class="sxs-lookup"><span data-stu-id="f5953-264">Requirements</span></span>

|<span data-ttu-id="f5953-265">要件</span><span class="sxs-lookup"><span data-stu-id="f5953-265">Requirement</span></span>| <span data-ttu-id="f5953-266">値</span><span class="sxs-lookup"><span data-stu-id="f5953-266">Value</span></span>|
|---|---|
|[<span data-ttu-id="f5953-267">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f5953-267">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f5953-268">1.0</span><span class="sxs-lookup"><span data-stu-id="f5953-268">1.0</span></span>|
|[<span data-ttu-id="f5953-269">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f5953-269">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f5953-270">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f5953-270">ReadItem</span></span>|
|[<span data-ttu-id="f5953-271">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f5953-271">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f5953-272">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="f5953-272">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f5953-273">戻り値:</span><span class="sxs-lookup"><span data-stu-id="f5953-273">Returns:</span></span>

<span data-ttu-id="f5953-274">型:[LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="f5953-274">Type: [LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)</span></span>

####  <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="f5953-275">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="f5953-275">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="f5953-276">EWS 形式のアイテム ID を REST 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="f5953-276">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="f5953-277">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f5953-277">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="f5953-p110">EWS または `itemId` プロパティで取得されるアイテム ID は、REST API ([Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) や [Microsoft Graph](https://graph.microsoft.io/) など) に使用される形式とは異なる形式を使用します。`convertToRestId` メソッドは、EWS 形式の ID を REST 用の適切な形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="f5953-p110">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f5953-280">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="f5953-280">Parameters:</span></span>

|<span data-ttu-id="f5953-281">名前</span><span class="sxs-lookup"><span data-stu-id="f5953-281">Name</span></span>| <span data-ttu-id="f5953-282">型</span><span class="sxs-lookup"><span data-stu-id="f5953-282">Type</span></span>| <span data-ttu-id="f5953-283">説明</span><span class="sxs-lookup"><span data-stu-id="f5953-283">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="f5953-284">String</span><span class="sxs-lookup"><span data-stu-id="f5953-284">String</span></span>|<span data-ttu-id="f5953-285">Exchange Web サービス (EWS) 形式のアイテム ID</span><span class="sxs-lookup"><span data-stu-id="f5953-285">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="f5953-286">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="f5953-286">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.restversion)|<span data-ttu-id="f5953-287">変換後の ID を使用する Outlook REST API のバージョンを示す値。</span><span class="sxs-lookup"><span data-stu-id="f5953-287">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f5953-288">要件</span><span class="sxs-lookup"><span data-stu-id="f5953-288">Requirements</span></span>

|<span data-ttu-id="f5953-289">要件</span><span class="sxs-lookup"><span data-stu-id="f5953-289">Requirement</span></span>| <span data-ttu-id="f5953-290">値</span><span class="sxs-lookup"><span data-stu-id="f5953-290">Value</span></span>|
|---|---|
|[<span data-ttu-id="f5953-291">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f5953-291">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f5953-292">1.3</span><span class="sxs-lookup"><span data-stu-id="f5953-292">1.3</span></span>|
|[<span data-ttu-id="f5953-293">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f5953-293">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f5953-294">制限あり</span><span class="sxs-lookup"><span data-stu-id="f5953-294">Restricted</span></span>|
|[<span data-ttu-id="f5953-295">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f5953-295">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f5953-296">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="f5953-296">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f5953-297">戻り値:</span><span class="sxs-lookup"><span data-stu-id="f5953-297">Returns:</span></span>

<span data-ttu-id="f5953-298">型:String</span><span class="sxs-lookup"><span data-stu-id="f5953-298">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="f5953-299">例</span><span class="sxs-lookup"><span data-stu-id="f5953-299">Example</span></span>

```js
// Get the currently selected item's ID
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the
// Outlook Mail API
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="f5953-300">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="f5953-300">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="f5953-301">時間情報が含まれているディクショナリから日付オブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="f5953-301">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="f5953-302">`convertToUtcClientTime` メソッドは、ローカルの日付と時刻を含むディクショナリを、ローカルの日付と時刻の正しい値を持つ日付オブジェクトに変換します。</span><span class="sxs-lookup"><span data-stu-id="f5953-302">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f5953-303">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="f5953-303">Parameters:</span></span>

|<span data-ttu-id="f5953-304">名前</span><span class="sxs-lookup"><span data-stu-id="f5953-304">Name</span></span>| <span data-ttu-id="f5953-305">型</span><span class="sxs-lookup"><span data-stu-id="f5953-305">Type</span></span>| <span data-ttu-id="f5953-306">説明</span><span class="sxs-lookup"><span data-stu-id="f5953-306">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="f5953-307">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="f5953-307">LocalClientTime</span></span>](/javascript/api/outlook_1_6/office.LocalClientTime)|<span data-ttu-id="f5953-308">変換するローカル時刻の値。</span><span class="sxs-lookup"><span data-stu-id="f5953-308">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f5953-309">要件</span><span class="sxs-lookup"><span data-stu-id="f5953-309">Requirements</span></span>

|<span data-ttu-id="f5953-310">要件</span><span class="sxs-lookup"><span data-stu-id="f5953-310">Requirement</span></span>| <span data-ttu-id="f5953-311">値</span><span class="sxs-lookup"><span data-stu-id="f5953-311">Value</span></span>|
|---|---|
|[<span data-ttu-id="f5953-312">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f5953-312">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f5953-313">1.0</span><span class="sxs-lookup"><span data-stu-id="f5953-313">1.0</span></span>|
|[<span data-ttu-id="f5953-314">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f5953-314">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f5953-315">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f5953-315">ReadItem</span></span>|
|[<span data-ttu-id="f5953-316">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f5953-316">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f5953-317">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="f5953-317">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f5953-318">戻り値:</span><span class="sxs-lookup"><span data-stu-id="f5953-318">Returns:</span></span>

<span data-ttu-id="f5953-319">時間が UTC で表現された日付オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="f5953-319">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="f5953-320">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="f5953-320">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="f5953-321">Date</span><span class="sxs-lookup"><span data-stu-id="f5953-321">Date</span></span></dd>

</dl>

####  <a name="displayappointmentformitemid"></a><span data-ttu-id="f5953-322">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="f5953-322">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="f5953-323">既存の予定を表示します。</span><span class="sxs-lookup"><span data-stu-id="f5953-323">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="f5953-324">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f5953-324">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="f5953-325">`displayAppointmentForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存の予定を開きます。</span><span class="sxs-lookup"><span data-stu-id="f5953-325">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="f5953-p111">Outlook for Mac では、この方法を使って、定期的な系列の一部ではない単一の予定や定期的な系列のマスター予定を表示できます。ただし、系列のインスタンスは表示できません。これは、Outlook for Mac においては定期的な系列のインスタンスのプロパティ (アイテム ID を含む) にアクセスできないためです。</span><span class="sxs-lookup"><span data-stu-id="f5953-p111">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="f5953-328">Outlook Web App では、このメソッドは指定されたフォームの本文が 32 KB 以下の文字数の場合にフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="f5953-328">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="f5953-329">指定のアイテム識別子が既存の予定を表していない場合は、クライアント コンピューターまたはデバイスで空のウィンドウが開き、エラー メッセージは返されません。</span><span class="sxs-lookup"><span data-stu-id="f5953-329">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f5953-330">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="f5953-330">Parameters:</span></span>

|<span data-ttu-id="f5953-331">名前</span><span class="sxs-lookup"><span data-stu-id="f5953-331">Name</span></span>| <span data-ttu-id="f5953-332">型</span><span class="sxs-lookup"><span data-stu-id="f5953-332">Type</span></span>| <span data-ttu-id="f5953-333">説明</span><span class="sxs-lookup"><span data-stu-id="f5953-333">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="f5953-334">String</span><span class="sxs-lookup"><span data-stu-id="f5953-334">String</span></span>|<span data-ttu-id="f5953-335">既存の予定の Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="f5953-335">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f5953-336">要件</span><span class="sxs-lookup"><span data-stu-id="f5953-336">Requirements</span></span>

|<span data-ttu-id="f5953-337">要件</span><span class="sxs-lookup"><span data-stu-id="f5953-337">Requirement</span></span>| <span data-ttu-id="f5953-338">値</span><span class="sxs-lookup"><span data-stu-id="f5953-338">Value</span></span>|
|---|---|
|[<span data-ttu-id="f5953-339">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f5953-339">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f5953-340">1.0</span><span class="sxs-lookup"><span data-stu-id="f5953-340">1.0</span></span>|
|[<span data-ttu-id="f5953-341">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f5953-341">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f5953-342">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f5953-342">ReadItem</span></span>|
|[<span data-ttu-id="f5953-343">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f5953-343">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f5953-344">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="f5953-344">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f5953-345">例</span><span class="sxs-lookup"><span data-stu-id="f5953-345">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  <a name="displaymessageformitemid"></a><span data-ttu-id="f5953-346">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="f5953-346">displayMessageForm(itemId)</span></span>

<span data-ttu-id="f5953-347">既存のメッセージを表示します。</span><span class="sxs-lookup"><span data-stu-id="f5953-347">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="f5953-348">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f5953-348">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="f5953-349">`displayMessageForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存のメッセージを開きます。</span><span class="sxs-lookup"><span data-stu-id="f5953-349">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="f5953-350">Outlook Web App では、このメソッドは指定されたフォームの本文が 32 KB 以下の文字数の場合にフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="f5953-350">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="f5953-351">指定のアイテム識別子が既存のメッセージを表していない場合は、クライアント コンピューターにはメッセージは表示されず、エラー メッセージも返されません。</span><span class="sxs-lookup"><span data-stu-id="f5953-351">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="f5953-p112">予定を表す `itemId` を含む `displayMessageForm` を使用しないでください。既存の予定を表示するには、`displayAppointmentForm` メソッドを使用します。新しい予定を作成するフォームを表示するには、`displayNewAppointmentForm` メソッドを使用します。</span><span class="sxs-lookup"><span data-stu-id="f5953-p112">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f5953-354">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="f5953-354">Parameters:</span></span>

|<span data-ttu-id="f5953-355">名前</span><span class="sxs-lookup"><span data-stu-id="f5953-355">Name</span></span>| <span data-ttu-id="f5953-356">型</span><span class="sxs-lookup"><span data-stu-id="f5953-356">Type</span></span>| <span data-ttu-id="f5953-357">説明</span><span class="sxs-lookup"><span data-stu-id="f5953-357">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="f5953-358">String</span><span class="sxs-lookup"><span data-stu-id="f5953-358">String</span></span>|<span data-ttu-id="f5953-359">既存のメッセージの Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="f5953-359">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f5953-360">要件</span><span class="sxs-lookup"><span data-stu-id="f5953-360">Requirements</span></span>

|<span data-ttu-id="f5953-361">要件</span><span class="sxs-lookup"><span data-stu-id="f5953-361">Requirement</span></span>| <span data-ttu-id="f5953-362">値</span><span class="sxs-lookup"><span data-stu-id="f5953-362">Value</span></span>|
|---|---|
|[<span data-ttu-id="f5953-363">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f5953-363">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f5953-364">1.0</span><span class="sxs-lookup"><span data-stu-id="f5953-364">1.0</span></span>|
|[<span data-ttu-id="f5953-365">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f5953-365">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f5953-366">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f5953-366">ReadItem</span></span>|
|[<span data-ttu-id="f5953-367">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f5953-367">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f5953-368">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="f5953-368">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f5953-369">例</span><span class="sxs-lookup"><span data-stu-id="f5953-369">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="f5953-370">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="f5953-370">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="f5953-371">新しい予定を作成するためのフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="f5953-371">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="f5953-372">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f5953-372">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="f5953-p113">`displayNewAppointmentForm` メソッドを使用すると、ユーザーが新しい予定または会議を作成できるフォームが開きます。パラメーターを指定すると、予定のフォーム フィールドにパラメーターの内容が自動的に設定されます。</span><span class="sxs-lookup"><span data-stu-id="f5953-p113">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="f5953-p114">このメソッドは、Outlook Web App と OWA for Devices において、出席者フィールドが含まれるフォームを必ず表示します。入力引数として出席者を指定しないと、このメソッドにより **[保存]** ボタンのあるフォームが表示されます。出席者を指定した場合には、フォームにその出席者と **[送信]** ボタンが表示されます。</span><span class="sxs-lookup"><span data-stu-id="f5953-p114">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="f5953-p115">Outlook リッチ クライアントと Outlook RT で、`requiredAttendees`、`optionalAttendees`、または `resources` パラメーターに出席者またはリソースを指定し、このメソッドを実行すると、**[送信]** ボタンがある会議フォームが表示されます。受信者を指定せずにこのメソッドを実行すると、**[保存して閉じる]** ボタンがある予定フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="f5953-p115">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="f5953-380">パラメーターのいずれかが指定のサイズ制限を超える場合、または不明なパラメーター名が指定されている場合は、例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="f5953-380">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f5953-381">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="f5953-381">Parameters:</span></span>

> [!NOTE]
> <span data-ttu-id="f5953-382">すべてのパラメーターは省略可能です。</span><span class="sxs-lookup"><span data-stu-id="f5953-382">All parameters are optional.</span></span>

|<span data-ttu-id="f5953-383">名前</span><span class="sxs-lookup"><span data-stu-id="f5953-383">Name</span></span>| <span data-ttu-id="f5953-384">型</span><span class="sxs-lookup"><span data-stu-id="f5953-384">Type</span></span>| <span data-ttu-id="f5953-385">説明</span><span class="sxs-lookup"><span data-stu-id="f5953-385">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="f5953-386">Object</span><span class="sxs-lookup"><span data-stu-id="f5953-386">Object</span></span> | <span data-ttu-id="f5953-387">新しい予定を記述するパラメーターのディクショナリ。</span><span class="sxs-lookup"><span data-stu-id="f5953-387">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="f5953-388">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="f5953-388">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="f5953-p116">予定に必要な各出席者について、メール アドレスを含む文字列の配列、または `EmailAddressDetails` オブジェクトを含む配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="f5953-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="f5953-391">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="f5953-391">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="f5953-p117">予定の各任意出席者について、メール アドレスを含む文字列の配列、または `EmailAddressDetails` オブジェクトを含む配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="f5953-p117">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="f5953-394">日付</span><span class="sxs-lookup"><span data-stu-id="f5953-394">Date</span></span> | <span data-ttu-id="f5953-395">予定の開始日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="f5953-395">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="f5953-396">Date</span><span class="sxs-lookup"><span data-stu-id="f5953-396">Date</span></span> | <span data-ttu-id="f5953-397">予定の終了日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="f5953-397">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="f5953-398">String</span><span class="sxs-lookup"><span data-stu-id="f5953-398">String</span></span> | <span data-ttu-id="f5953-p118">予定の場所を含む文字列。文字列は最大 255 文字に制限されます。</span><span class="sxs-lookup"><span data-stu-id="f5953-p118">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="f5953-401">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="f5953-401">Array.&lt;String&gt;</span></span> | <span data-ttu-id="f5953-p119">予定に必要なリソースを含む文字列の配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="f5953-p119">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="f5953-404">String</span><span class="sxs-lookup"><span data-stu-id="f5953-404">String</span></span> | <span data-ttu-id="f5953-p120">予定の件名を含む文字列です。文字列は最大 255 文字に制限されます。</span><span class="sxs-lookup"><span data-stu-id="f5953-p120">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="f5953-407">String</span><span class="sxs-lookup"><span data-stu-id="f5953-407">String</span></span> | <span data-ttu-id="f5953-p121">予定の本文。本文の内容は、最大サイズが 32 KB に制限されます。</span><span class="sxs-lookup"><span data-stu-id="f5953-p121">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f5953-410">要件</span><span class="sxs-lookup"><span data-stu-id="f5953-410">Requirements</span></span>

|<span data-ttu-id="f5953-411">要件</span><span class="sxs-lookup"><span data-stu-id="f5953-411">Requirement</span></span>| <span data-ttu-id="f5953-412">値</span><span class="sxs-lookup"><span data-stu-id="f5953-412">Value</span></span>|
|---|---|
|[<span data-ttu-id="f5953-413">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f5953-413">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f5953-414">1.0</span><span class="sxs-lookup"><span data-stu-id="f5953-414">1.0</span></span>|
|[<span data-ttu-id="f5953-415">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f5953-415">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f5953-416">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f5953-416">ReadItem</span></span>|
|[<span data-ttu-id="f5953-417">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f5953-417">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f5953-418">読み取り</span><span class="sxs-lookup"><span data-stu-id="f5953-418">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f5953-419">例</span><span class="sxs-lookup"><span data-stu-id="f5953-419">Example</span></span>

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

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="f5953-420">displayNewMessageForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="f5953-420">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="f5953-421">新しいメッセージを作成するためのフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="f5953-421">Displays a form for creating a new message.</span></span>

<span data-ttu-id="f5953-422">`displayNewMessageForm` メソッドは、ユーザーが新しいメッセージを作成できるようにするフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="f5953-422">The `displayNewMessageForm` method opens a form that enables the user to create a new message.</span></span> <span data-ttu-id="f5953-423">パラメーターを指定すると、メッセージ フォーム フィールドにはパラメーターのコンテンツが自動的に入力されます。</span><span class="sxs-lookup"><span data-stu-id="f5953-423">If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="f5953-424">パラメーターのいずれかが指定のサイズ制限を超える場合、または不明なパラメーター名が指定されている場合は、例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="f5953-424">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f5953-425">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="f5953-425">Parameters:</span></span>

> [!NOTE]
> <span data-ttu-id="f5953-426">すべてのパラメーターは省略可能です。</span><span class="sxs-lookup"><span data-stu-id="f5953-426">All parameters are optional.</span></span>

|<span data-ttu-id="f5953-427">名前</span><span class="sxs-lookup"><span data-stu-id="f5953-427">Name</span></span>| <span data-ttu-id="f5953-428">型</span><span class="sxs-lookup"><span data-stu-id="f5953-428">Type</span></span>| <span data-ttu-id="f5953-429">説明</span><span class="sxs-lookup"><span data-stu-id="f5953-429">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="f5953-430">Object</span><span class="sxs-lookup"><span data-stu-id="f5953-430">Object</span></span> | <span data-ttu-id="f5953-431">新しいメッセージを記述するパラメーターのディクショナリ。</span><span class="sxs-lookup"><span data-stu-id="f5953-431">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="f5953-432">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="f5953-432">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="f5953-433">メール アドレスを含む文字列の配列、または To 行の各受信者の `EmailAddressDetails` オブジェクトを含む配列。</span><span class="sxs-lookup"><span data-stu-id="f5953-433">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line.</span></span> <span data-ttu-id="f5953-434">配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="f5953-434">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="f5953-435">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="f5953-435">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="f5953-436">メール アドレスを含む文字列の配列、または Cc 行の各受信者の `EmailAddressDetails` オブジェクトを含む配列。</span><span class="sxs-lookup"><span data-stu-id="f5953-436">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line.</span></span> <span data-ttu-id="f5953-437">配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="f5953-437">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="f5953-438">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="f5953-438">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="f5953-439">メール アドレスを含む文字列の配列、または Bcc 行の各受信者の `EmailAddressDetails` オブジェクトを含む配列。</span><span class="sxs-lookup"><span data-stu-id="f5953-439">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line.</span></span> <span data-ttu-id="f5953-440">配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="f5953-440">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="f5953-441">String</span><span class="sxs-lookup"><span data-stu-id="f5953-441">String</span></span> | <span data-ttu-id="f5953-442">メッセージの件名を含む文字列。</span><span class="sxs-lookup"><span data-stu-id="f5953-442">A string containing the subject of the message.</span></span> <span data-ttu-id="f5953-443">文字列は最大 255 文字に制限されます。</span><span class="sxs-lookup"><span data-stu-id="f5953-443">The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="f5953-444">String</span><span class="sxs-lookup"><span data-stu-id="f5953-444">String</span></span> | <span data-ttu-id="f5953-445">メッセージの HTML 本文。</span><span class="sxs-lookup"><span data-stu-id="f5953-445">The HTML body of the message.</span></span> <span data-ttu-id="f5953-446">本文の内容は、最大サイズが 32 KB に制限されます。</span><span class="sxs-lookup"><span data-stu-id="f5953-446">The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="f5953-447">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="f5953-447">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="f5953-448">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="f5953-448">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="f5953-449">String</span><span class="sxs-lookup"><span data-stu-id="f5953-449">String</span></span> | <span data-ttu-id="f5953-p128">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="f5953-p128">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="f5953-452">String</span><span class="sxs-lookup"><span data-stu-id="f5953-452">String</span></span> | <span data-ttu-id="f5953-453">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="f5953-453">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="f5953-454">String</span><span class="sxs-lookup"><span data-stu-id="f5953-454">String</span></span> | <span data-ttu-id="f5953-p129">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="f5953-p129">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="f5953-457">ブール値</span><span class="sxs-lookup"><span data-stu-id="f5953-457">Boolean</span></span> | <span data-ttu-id="f5953-p130">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="f5953-p130">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="f5953-460">String</span><span class="sxs-lookup"><span data-stu-id="f5953-460">String</span></span> | <span data-ttu-id="f5953-461">`type` が `item` に設定されている場合にのみ使用されます。</span><span class="sxs-lookup"><span data-stu-id="f5953-461">Only used if `type` is set to `item`.</span></span> <span data-ttu-id="f5953-462">新しいメッセージに添付する必要がある既存の電子メールの EWS のアイテム ID です。</span><span class="sxs-lookup"><span data-stu-id="f5953-462">The EWS item id of the existing e-mail you want to attach to the new message.</span></span> <span data-ttu-id="f5953-463">最大 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="f5953-463">This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="f5953-464">要件</span><span class="sxs-lookup"><span data-stu-id="f5953-464">Requirements</span></span>

|<span data-ttu-id="f5953-465">要件</span><span class="sxs-lookup"><span data-stu-id="f5953-465">Requirement</span></span>| <span data-ttu-id="f5953-466">値</span><span class="sxs-lookup"><span data-stu-id="f5953-466">Value</span></span>|
|---|---|
|[<span data-ttu-id="f5953-467">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f5953-467">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f5953-468">1.6</span><span class="sxs-lookup"><span data-stu-id="f5953-468">1.6</span></span> |
|[<span data-ttu-id="f5953-469">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f5953-469">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f5953-470">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f5953-470">ReadItem</span></span>|
|[<span data-ttu-id="f5953-471">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f5953-471">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f5953-472">読み取り</span><span class="sxs-lookup"><span data-stu-id="f5953-472">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f5953-473">例</span><span class="sxs-lookup"><span data-stu-id="f5953-473">Example</span></span>

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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="f5953-474">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="f5953-474">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="f5953-475">REST API または Exchange Web サービスを呼び出すために使用するトークンを含む文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="f5953-475">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="f5953-p132">`getCallbackTokenAsync` メソッドは、ユーザーのメールボックスをホストする Exchange Server から不透明なトークンを取得する非同期の呼び出しを行います。コールバック トークンの有効期間は 5 分です。</span><span class="sxs-lookup"><span data-stu-id="f5953-p132">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="f5953-478">可能であれば、アドインでは Exchange Web サービスの代わりに REST API を使用することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="f5953-478">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span> 

<span data-ttu-id="f5953-479">**REST トークン**</span><span class="sxs-lookup"><span data-stu-id="f5953-479">**REST Tokens**</span></span>

<span data-ttu-id="f5953-p133">REST トークンが要求された場合 (`options.isRest = true`)、結果トークンは Exchange Web サービスの呼び出しを認証するためには機能しません。アドインがマニフェストで [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) アクセス許可を指定していない限り、トークンの範囲は現在のアイテムとその添付ファイルへの読み取り専用アクセスに制限されます。`ReadWriteMailbox` アクセス許可が指定されている場合は、結果トークンは、メールを送信する機能など、メール、カレンダー、連絡先への読み取り/書き込みアクセスを付与します。</span><span class="sxs-lookup"><span data-stu-id="f5953-p133">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="f5953-483">アドインでは、`restUrl` プロパティを使用して、REST API 呼び出しを行うときに使用する正しい URL を決定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="f5953-483">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="f5953-484">**EWS トークン**</span><span class="sxs-lookup"><span data-stu-id="f5953-484">**EWS Tokens**</span></span>

<span data-ttu-id="f5953-p134">EWS トークンが要求された場合 (`options.isRest = false`)、結果トークンは REST API 呼び出しを認証するためには機能しません。トークンの範囲は、現在のアイテムへのアクセスに制限されます。</span><span class="sxs-lookup"><span data-stu-id="f5953-p134">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="f5953-487">アドインでは、`ewsUrl` プロパティを使用して、EWS 呼び出しを行うときに使用する正しい URL を決定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="f5953-487">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f5953-488">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="f5953-488">Parameters:</span></span>

|<span data-ttu-id="f5953-489">名前</span><span class="sxs-lookup"><span data-stu-id="f5953-489">Name</span></span>| <span data-ttu-id="f5953-490">型</span><span class="sxs-lookup"><span data-stu-id="f5953-490">Type</span></span>| <span data-ttu-id="f5953-491">属性</span><span class="sxs-lookup"><span data-stu-id="f5953-491">Attributes</span></span>| <span data-ttu-id="f5953-492">説明</span><span class="sxs-lookup"><span data-stu-id="f5953-492">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="f5953-493">Object</span><span class="sxs-lookup"><span data-stu-id="f5953-493">Object</span></span> | <span data-ttu-id="f5953-494">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f5953-494">&lt;optional&gt;</span></span> | <span data-ttu-id="f5953-495">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="f5953-495">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="f5953-496">Boolean</span><span class="sxs-lookup"><span data-stu-id="f5953-496">Boolean</span></span> |  <span data-ttu-id="f5953-497">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="f5953-497">&lt;optional&gt;</span></span> | <span data-ttu-id="f5953-p135">提供されたトークンを Outlook REST API または Exchange Web サービスに使用するかどうかを決定します。既定値は、`false` です。</span><span class="sxs-lookup"><span data-stu-id="f5953-p135">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="f5953-500">Object</span><span class="sxs-lookup"><span data-stu-id="f5953-500">Object</span></span> |  <span data-ttu-id="f5953-501">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="f5953-501">&lt;optional&gt;</span></span> | <span data-ttu-id="f5953-502">非同期メソッドに渡される状態データ。</span><span class="sxs-lookup"><span data-stu-id="f5953-502">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="f5953-503">function</span><span class="sxs-lookup"><span data-stu-id="f5953-503">function</span></span>||<span data-ttu-id="f5953-p136">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="f5953-p136">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f5953-506">要件</span><span class="sxs-lookup"><span data-stu-id="f5953-506">Requirements</span></span>

|<span data-ttu-id="f5953-507">要件</span><span class="sxs-lookup"><span data-stu-id="f5953-507">Requirement</span></span>| <span data-ttu-id="f5953-508">値</span><span class="sxs-lookup"><span data-stu-id="f5953-508">Value</span></span>|
|---|---|
|[<span data-ttu-id="f5953-509">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f5953-509">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f5953-510">1.5</span><span class="sxs-lookup"><span data-stu-id="f5953-510">1.5</span></span> |
|[<span data-ttu-id="f5953-511">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f5953-511">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f5953-512">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f5953-512">ReadItem</span></span>|
|[<span data-ttu-id="f5953-513">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f5953-513">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f5953-514">新規作成と閲覧</span><span class="sxs-lookup"><span data-stu-id="f5953-514">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="f5953-515">例</span><span class="sxs-lookup"><span data-stu-id="f5953-515">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="f5953-516">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="f5953-516">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="f5953-517">Exchange Server から添付ファイルやアイテムを取得するために使うトークンを含む文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="f5953-517">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="f5953-p137">`getCallbackTokenAsync` メソッドは、ユーザーのメールボックスをホストする Exchange Server から不透明なトークンを取得する非同期の呼び出しを行います。コールバック トークンの有効期間は 5 分です。</span><span class="sxs-lookup"><span data-stu-id="f5953-p137">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="f5953-p138">トークンと予定の識別子またはアイテムの識別子をサードパーティ システムに渡すことができます。サードパーティ システムは、トークンをベアラー承認トークンとして使用し、Exchange Web サービス (EWS) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) または [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) 操作を呼び出して、添付ファイルまたはアイテムを返します。たとえば、[選択したアイテムから添付ファイルを取得する](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="f5953-p138">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="f5953-523">アプリが閲覧モードで `getCallbackTokenAsync` メソッドを呼び出すには、アプリのマニフェスト内に **ReadItem** アクセス許可が指定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="f5953-523">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="f5953-p139">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出してアイテムの識別子を `getCallbackTokenAsync` メソッドに渡す必要があります。アプリには、`saveAsync` メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="f5953-p139">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f5953-526">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="f5953-526">Parameters:</span></span>

|<span data-ttu-id="f5953-527">名前</span><span class="sxs-lookup"><span data-stu-id="f5953-527">Name</span></span>| <span data-ttu-id="f5953-528">型</span><span class="sxs-lookup"><span data-stu-id="f5953-528">Type</span></span>| <span data-ttu-id="f5953-529">属性</span><span class="sxs-lookup"><span data-stu-id="f5953-529">Attributes</span></span>| <span data-ttu-id="f5953-530">説明</span><span class="sxs-lookup"><span data-stu-id="f5953-530">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="f5953-531">function</span><span class="sxs-lookup"><span data-stu-id="f5953-531">function</span></span>||<span data-ttu-id="f5953-p140">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="f5953-p140">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="f5953-534">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="f5953-534">Object</span></span>| <span data-ttu-id="f5953-535">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="f5953-535">&lt;optional&gt;</span></span>|<span data-ttu-id="f5953-536">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="f5953-536">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f5953-537">要件</span><span class="sxs-lookup"><span data-stu-id="f5953-537">Requirements</span></span>

|<span data-ttu-id="f5953-538">要件</span><span class="sxs-lookup"><span data-stu-id="f5953-538">Requirement</span></span>| <span data-ttu-id="f5953-539">値</span><span class="sxs-lookup"><span data-stu-id="f5953-539">Value</span></span>|
|---|---|
|[<span data-ttu-id="f5953-540">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f5953-540">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f5953-541">1.3</span><span class="sxs-lookup"><span data-stu-id="f5953-541">1.3</span></span>|
|[<span data-ttu-id="f5953-542">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f5953-542">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f5953-543">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f5953-543">ReadItem</span></span>|
|[<span data-ttu-id="f5953-544">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f5953-544">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f5953-545">新規作成と閲覧</span><span class="sxs-lookup"><span data-stu-id="f5953-545">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="f5953-546">例</span><span class="sxs-lookup"><span data-stu-id="f5953-546">Example</span></span>

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="f5953-547">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="f5953-547">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="f5953-548">ユーザーと Office アドインを識別するトークンを取得します。</span><span class="sxs-lookup"><span data-stu-id="f5953-548">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="f5953-549">`getUserIdentityTokenAsync` メソッドは、[アドインとユーザーをサード パーティのシステムで識別して認証](https://docs.microsoft.com/outlook/add-ins/authentication)することのできるトークンを返します。</span><span class="sxs-lookup"><span data-stu-id="f5953-549">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](https://docs.microsoft.com/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="f5953-550">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="f5953-550">Parameters:</span></span>

|<span data-ttu-id="f5953-551">名前</span><span class="sxs-lookup"><span data-stu-id="f5953-551">Name</span></span>| <span data-ttu-id="f5953-552">型</span><span class="sxs-lookup"><span data-stu-id="f5953-552">Type</span></span>| <span data-ttu-id="f5953-553">属性</span><span class="sxs-lookup"><span data-stu-id="f5953-553">Attributes</span></span>| <span data-ttu-id="f5953-554">説明</span><span class="sxs-lookup"><span data-stu-id="f5953-554">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="f5953-555">function</span><span class="sxs-lookup"><span data-stu-id="f5953-555">function</span></span>||<span data-ttu-id="f5953-556">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="f5953-556">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="f5953-557">トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="f5953-557">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="f5953-558">Object</span><span class="sxs-lookup"><span data-stu-id="f5953-558">Object</span></span>| <span data-ttu-id="f5953-559">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="f5953-559">&lt;optional&gt;</span></span>|<span data-ttu-id="f5953-560">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="f5953-560">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f5953-561">要件</span><span class="sxs-lookup"><span data-stu-id="f5953-561">Requirements</span></span>

|<span data-ttu-id="f5953-562">要件</span><span class="sxs-lookup"><span data-stu-id="f5953-562">Requirement</span></span>| <span data-ttu-id="f5953-563">値</span><span class="sxs-lookup"><span data-stu-id="f5953-563">Value</span></span>|
|---|---|
|[<span data-ttu-id="f5953-564">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f5953-564">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f5953-565">1.0</span><span class="sxs-lookup"><span data-stu-id="f5953-565">1.0</span></span>|
|[<span data-ttu-id="f5953-566">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f5953-566">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f5953-567">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f5953-567">ReadItem</span></span>|
|[<span data-ttu-id="f5953-568">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f5953-568">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f5953-569">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="f5953-569">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f5953-570">例</span><span class="sxs-lookup"><span data-stu-id="f5953-570">Example</span></span>

```js
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="f5953-571">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="f5953-571">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="f5953-572">ユーザーのメールボックスをホストしている Exchange サーバー上の Exchange Web サービス (EWS) のサービスに対して非同期の要求を行います。</span><span class="sxs-lookup"><span data-stu-id="f5953-572">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="f5953-573">このメソッドは、次のシナリオではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f5953-573">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="f5953-574">Outlook for iOS または Outlook for Android を使用している場合</span><span class="sxs-lookup"><span data-stu-id="f5953-574">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="f5953-575">アドインが Gmail のメールボックスに読み込まれる場合</span><span class="sxs-lookup"><span data-stu-id="f5953-575">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="f5953-576">このような場合は、アドインでは [REST API を使用](https://docs.microsoft.com/outlook/add-ins/use-rest-api)して、代わりにユーザーのメールボックスにアクセスする必要があります。</span><span class="sxs-lookup"><span data-stu-id="f5953-576">In these cases, add-ins should [use REST APIs](https://docs.microsoft.com/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="f5953-577">`makeEwsRequestAsync` メソッドは、アドインの代わりに Exchange に EWS 要求を送信します。</span><span class="sxs-lookup"><span data-stu-id="f5953-577">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="f5953-578">サポートされている EWS 操作の一覧については、「[Outlook アドインからの Web サービスの呼び出し](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f5953-578">See [Call web services from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="f5953-579">`makeEwsRequestAsync` メソッドでは、フォルダー関連アイテムを要求できません。</span><span class="sxs-lookup"><span data-stu-id="f5953-579">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="f5953-580">XML 要求では UTF-8 エンコードを指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="f5953-580">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="f5953-p142">`makeEwsRequestAsync` メソッドを使用するには、アドインに **ReadWriteMailbox** アクセス許可が必要です。**ReadWriteMailbox** アクセス許可と、`makeEwsRequestAsync` メソッドで呼び出せる EWS 操作の使い方については、「[ユーザーのメールボックスへのメール アドイン アクセスのアクセス許可を指定する](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f5953-p142">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="f5953-583">サーバー管理者は、クライアント アクセス サーバーの EWS ディレクトリで `OAuthAuthentication` を true に設定して、`makeEwsRequestAsync` メソッドで EWS 要求を行うことができるようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="f5953-583">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="f5953-584">バージョンの相違点</span><span class="sxs-lookup"><span data-stu-id="f5953-584">Version differences</span></span>

<span data-ttu-id="f5953-585">バージョン 15.0.4535.1004 より前のバージョンの Outlook で実行しているメール アプリで `makeEwsRequestAsync` メソッドを使う場合は、エンコード値を `ISO-8859-1` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="f5953-585">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="f5953-p143">Outlook on the web でメール アプリを実行している場合は、エンコード値を設定する必要はありません。mailbox.diagnostics.hostName プロパティを使って、メール アプリを Outlook で実行しているのか、Outlook on the web で実行しているのかを確認できます。mailbox.diagnostics.hostVersion プロパティを使って、どのバージョンの Outlook を使って実行しているかを確認できます。</span><span class="sxs-lookup"><span data-stu-id="f5953-p143">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f5953-589">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="f5953-589">Parameters:</span></span>

|<span data-ttu-id="f5953-590">名前</span><span class="sxs-lookup"><span data-stu-id="f5953-590">Name</span></span>| <span data-ttu-id="f5953-591">型</span><span class="sxs-lookup"><span data-stu-id="f5953-591">Type</span></span>| <span data-ttu-id="f5953-592">属性</span><span class="sxs-lookup"><span data-stu-id="f5953-592">Attributes</span></span>| <span data-ttu-id="f5953-593">説明</span><span class="sxs-lookup"><span data-stu-id="f5953-593">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="f5953-594">String</span><span class="sxs-lookup"><span data-stu-id="f5953-594">String</span></span>||<span data-ttu-id="f5953-595">EWS 要求です。</span><span class="sxs-lookup"><span data-stu-id="f5953-595">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="f5953-596">function</span><span class="sxs-lookup"><span data-stu-id="f5953-596">function</span></span>||<span data-ttu-id="f5953-597">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="f5953-597">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="f5953-598">EWS 呼び出しの XML 結果は、`asyncResult.value` プロパティ内の文字列として提供されています。</span><span class="sxs-lookup"><span data-stu-id="f5953-598">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="f5953-599">結果のサイズが 1 MB を超える場合、代わりにエラー メッセージが返されます。</span><span class="sxs-lookup"><span data-stu-id="f5953-599">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="f5953-600">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="f5953-600">Object</span></span>| <span data-ttu-id="f5953-601">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="f5953-601">&lt;optional&gt;</span></span>|<span data-ttu-id="f5953-602">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="f5953-602">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f5953-603">要件</span><span class="sxs-lookup"><span data-stu-id="f5953-603">Requirements</span></span>

|<span data-ttu-id="f5953-604">要件</span><span class="sxs-lookup"><span data-stu-id="f5953-604">Requirement</span></span>| <span data-ttu-id="f5953-605">値</span><span class="sxs-lookup"><span data-stu-id="f5953-605">Value</span></span>|
|---|---|
|[<span data-ttu-id="f5953-606">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f5953-606">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f5953-607">1.0</span><span class="sxs-lookup"><span data-stu-id="f5953-607">1.0</span></span>|
|[<span data-ttu-id="f5953-608">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f5953-608">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f5953-609">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="f5953-609">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="f5953-610">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f5953-610">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f5953-611">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="f5953-611">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f5953-612">例</span><span class="sxs-lookup"><span data-stu-id="f5953-612">Example</span></span>

<span data-ttu-id="f5953-613">次の例は、`GetItem` 操作を使ってアイテムの件名を取得するため、`makeEwsRequestAsync` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="f5953-613">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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

####  <a name="removehandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="f5953-614">removeHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="f5953-614">removeHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="f5953-615">サポートされているイベントのイベント ハンドラーを削除します。</span><span class="sxs-lookup"><span data-stu-id="f5953-615">Removes an event handler for a supported event.</span></span>

<span data-ttu-id="f5953-616">現在、サポートされているイベントの種類は `Office.EventType.ItemChanged` だけです。</span><span class="sxs-lookup"><span data-stu-id="f5953-616">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f5953-617">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="f5953-617">Parameters:</span></span>

| <span data-ttu-id="f5953-618">名前</span><span class="sxs-lookup"><span data-stu-id="f5953-618">Name</span></span> | <span data-ttu-id="f5953-619">型</span><span class="sxs-lookup"><span data-stu-id="f5953-619">Type</span></span> | <span data-ttu-id="f5953-620">属性</span><span class="sxs-lookup"><span data-stu-id="f5953-620">Attributes</span></span> | <span data-ttu-id="f5953-621">説明</span><span class="sxs-lookup"><span data-stu-id="f5953-621">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="f5953-622">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="f5953-622">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="f5953-623">ハンドラーを取り消すイベント。</span><span class="sxs-lookup"><span data-stu-id="f5953-623">The event that should revoke the handler.</span></span> |
| `handler` | <span data-ttu-id="f5953-624">職務</span><span class="sxs-lookup"><span data-stu-id="f5953-624">Function</span></span> || <span data-ttu-id="f5953-p145">イベントを処理する関数。関数は、オブジェクト リテラルである単一パラメーターを受け入れる必要があります。パラメーターの `type` プロパティは、`addHandlerAsync` に渡される `eventType` パラメーターと一致します。</span><span class="sxs-lookup"><span data-stu-id="f5953-p145">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="f5953-628">Object</span><span class="sxs-lookup"><span data-stu-id="f5953-628">Object</span></span> | <span data-ttu-id="f5953-629">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f5953-629">&lt;optional&gt;</span></span> | <span data-ttu-id="f5953-630">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="f5953-630">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="f5953-631">Object</span><span class="sxs-lookup"><span data-stu-id="f5953-631">Object</span></span> | <span data-ttu-id="f5953-632">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f5953-632">&lt;optional&gt;</span></span> | <span data-ttu-id="f5953-633">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="f5953-633">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="f5953-634">function</span><span class="sxs-lookup"><span data-stu-id="f5953-634">function</span></span>| <span data-ttu-id="f5953-635">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f5953-635">&lt;optional&gt;</span></span>|<span data-ttu-id="f5953-636">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="f5953-636">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f5953-637">要件</span><span class="sxs-lookup"><span data-stu-id="f5953-637">Requirements</span></span>

|<span data-ttu-id="f5953-638">要件</span><span class="sxs-lookup"><span data-stu-id="f5953-638">Requirement</span></span>| <span data-ttu-id="f5953-639">値</span><span class="sxs-lookup"><span data-stu-id="f5953-639">Value</span></span>|
|---|---|
|[<span data-ttu-id="f5953-640">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f5953-640">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f5953-641">1.5</span><span class="sxs-lookup"><span data-stu-id="f5953-641">1.5</span></span> |
|[<span data-ttu-id="f5953-642">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f5953-642">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f5953-643">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f5953-643">ReadItem</span></span> |
|[<span data-ttu-id="f5953-644">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f5953-644">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f5953-645">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="f5953-645">Compose or read</span></span>|