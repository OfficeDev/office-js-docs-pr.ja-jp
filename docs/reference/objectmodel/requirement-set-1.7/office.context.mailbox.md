---
title: Office. メールボックス要件セット1.7
description: ''
ms.date: 08/30/2019
localization_priority: Normal
ms.openlocfilehash: 10165f68edee3f4ac0df1ff053d4e64fb009a766
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/30/2019
ms.locfileid: "36695960"
---
# <a name="mailbox"></a><span data-ttu-id="9f271-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="9f271-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="9f271-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="9f271-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="9f271-104">Microsoft Outlook の Outlook アドインオブジェクトモデルへのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="9f271-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="9f271-105">要件</span><span class="sxs-lookup"><span data-stu-id="9f271-105">Requirements</span></span>

|<span data-ttu-id="9f271-106">要件</span><span class="sxs-lookup"><span data-stu-id="9f271-106">Requirement</span></span>| <span data-ttu-id="9f271-107">値</span><span class="sxs-lookup"><span data-stu-id="9f271-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="9f271-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="9f271-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9f271-109">1.0</span><span class="sxs-lookup"><span data-stu-id="9f271-109">1.0</span></span>|
|[<span data-ttu-id="9f271-110">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="9f271-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9f271-111">制限あり</span><span class="sxs-lookup"><span data-stu-id="9f271-111">Restricted</span></span>|
|[<span data-ttu-id="9f271-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="9f271-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9f271-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="9f271-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="9f271-114">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="9f271-114">Members and methods</span></span>

| <span data-ttu-id="9f271-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="9f271-115">Member</span></span> | <span data-ttu-id="9f271-116">種類</span><span class="sxs-lookup"><span data-stu-id="9f271-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="9f271-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="9f271-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="9f271-118">メンバー</span><span class="sxs-lookup"><span data-stu-id="9f271-118">Member</span></span> |
| [<span data-ttu-id="9f271-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="9f271-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="9f271-120">メンバー</span><span class="sxs-lookup"><span data-stu-id="9f271-120">Member</span></span> |
| [<span data-ttu-id="9f271-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="9f271-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="9f271-122">メソッド</span><span class="sxs-lookup"><span data-stu-id="9f271-122">Method</span></span> |
| [<span data-ttu-id="9f271-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="9f271-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="9f271-124">メソッド</span><span class="sxs-lookup"><span data-stu-id="9f271-124">Method</span></span> |
| [<span data-ttu-id="9f271-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="9f271-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="9f271-126">メソッド</span><span class="sxs-lookup"><span data-stu-id="9f271-126">Method</span></span> |
| [<span data-ttu-id="9f271-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="9f271-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="9f271-128">メソッド</span><span class="sxs-lookup"><span data-stu-id="9f271-128">Method</span></span> |
| [<span data-ttu-id="9f271-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="9f271-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="9f271-130">メソッド</span><span class="sxs-lookup"><span data-stu-id="9f271-130">Method</span></span> |
| [<span data-ttu-id="9f271-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="9f271-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="9f271-132">メソッド</span><span class="sxs-lookup"><span data-stu-id="9f271-132">Method</span></span> |
| [<span data-ttu-id="9f271-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="9f271-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="9f271-134">メソッド</span><span class="sxs-lookup"><span data-stu-id="9f271-134">Method</span></span> |
| [<span data-ttu-id="9f271-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="9f271-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="9f271-136">メソッド</span><span class="sxs-lookup"><span data-stu-id="9f271-136">Method</span></span> |
| [<span data-ttu-id="9f271-137">displayNewMessageForm</span><span class="sxs-lookup"><span data-stu-id="9f271-137">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="9f271-138">メソッド</span><span class="sxs-lookup"><span data-stu-id="9f271-138">Method</span></span> |
| [<span data-ttu-id="9f271-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="9f271-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="9f271-140">メソッド</span><span class="sxs-lookup"><span data-stu-id="9f271-140">Method</span></span> |
| [<span data-ttu-id="9f271-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="9f271-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="9f271-142">メソッド</span><span class="sxs-lookup"><span data-stu-id="9f271-142">Method</span></span> |
| [<span data-ttu-id="9f271-143">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="9f271-143">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="9f271-144">メソッド</span><span class="sxs-lookup"><span data-stu-id="9f271-144">Method</span></span> |
| [<span data-ttu-id="9f271-145">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="9f271-145">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="9f271-146">メソッド</span><span class="sxs-lookup"><span data-stu-id="9f271-146">Method</span></span> |
| [<span data-ttu-id="9f271-147">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="9f271-147">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="9f271-148">メソッド</span><span class="sxs-lookup"><span data-stu-id="9f271-148">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="9f271-149">名前空間</span><span class="sxs-lookup"><span data-stu-id="9f271-149">Namespaces</span></span>

<span data-ttu-id="9f271-150">[diagnostics](Office.context.mailbox.diagnostics.md):Outlook アドインに診断情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="9f271-150">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="9f271-151">[item](Office.context.mailbox.item.md):Outlook アドインのメッセージや予定にアクセスするためのメソッドとプロパティを提供します。</span><span class="sxs-lookup"><span data-stu-id="9f271-151">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="9f271-152">[userProfile](Office.context.mailbox.userProfile.md):Outlook アドインのユーザーに関する情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="9f271-152">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="9f271-153">メンバー</span><span class="sxs-lookup"><span data-stu-id="9f271-153">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="9f271-154">ewsUrl: String</span><span class="sxs-lookup"><span data-stu-id="9f271-154">ewsUrl: String</span></span>

<span data-ttu-id="9f271-155">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account.</span><span class="sxs-lookup"><span data-stu-id="9f271-155">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account.</span></span> <span data-ttu-id="9f271-156">Read mode only.</span><span class="sxs-lookup"><span data-stu-id="9f271-156">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="9f271-157">このメンバーは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="9f271-157">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="9f271-p102">
  `ewsUrl\` 値は、リモート サービスで、ユーザーのメールボックスに EWS 呼び出しを行うために使うことができます。たとえば、[選択したアイテムから添付ファイルを取得する](/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="9f271-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="9f271-160">アプリが閲覧モードで `ewsUrl` メンバーを呼び出すには、アプリのマニフェスト内に **ReadItem** アクセス許可が指定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="9f271-160">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="9f271-p103">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出してから、`ewsUrl` メンバーを使用する必要があります。アプリには、`saveAsync` メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="9f271-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="9f271-163">型</span><span class="sxs-lookup"><span data-stu-id="9f271-163">Type</span></span>

*   <span data-ttu-id="9f271-164">String</span><span class="sxs-lookup"><span data-stu-id="9f271-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9f271-165">要件</span><span class="sxs-lookup"><span data-stu-id="9f271-165">Requirements</span></span>

|<span data-ttu-id="9f271-166">要件</span><span class="sxs-lookup"><span data-stu-id="9f271-166">Requirement</span></span>| <span data-ttu-id="9f271-167">値</span><span class="sxs-lookup"><span data-stu-id="9f271-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="9f271-168">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="9f271-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9f271-169">1.0</span><span class="sxs-lookup"><span data-stu-id="9f271-169">1.0</span></span>|
|[<span data-ttu-id="9f271-170">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="9f271-170">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9f271-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9f271-171">ReadItem</span></span>|
|[<span data-ttu-id="9f271-172">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="9f271-172">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9f271-173">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="9f271-173">Compose or Read</span></span>|

<br>

---
---

#### <a name="resturl-string"></a><span data-ttu-id="9f271-174">Office.context.mailbox.resturl が: String</span><span class="sxs-lookup"><span data-stu-id="9f271-174">restUrl: String</span></span>

<span data-ttu-id="9f271-175">この電子メール アカウントの REST エンドポイントの URL を取得します。</span><span class="sxs-lookup"><span data-stu-id="9f271-175">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="9f271-176">`restUrl` 値は、ユーザーのメールボックスに [REST API](/outlook/rest/) 呼び出しを行うために使用できます。</span><span class="sxs-lookup"><span data-stu-id="9f271-176">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="9f271-177">アプリが閲覧モードで `restUrl` メンバーを呼び出すには、アプリのマニフェスト内に **ReadItem** アクセス許可が指定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="9f271-177">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="9f271-p104">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出してから、`restUrl` メンバーを使用する必要があります。アプリには、`saveAsync` メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="9f271-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="9f271-180">型</span><span class="sxs-lookup"><span data-stu-id="9f271-180">Type</span></span>

*   <span data-ttu-id="9f271-181">String</span><span class="sxs-lookup"><span data-stu-id="9f271-181">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9f271-182">要件</span><span class="sxs-lookup"><span data-stu-id="9f271-182">Requirements</span></span>

|<span data-ttu-id="9f271-183">要件</span><span class="sxs-lookup"><span data-stu-id="9f271-183">Requirement</span></span>| <span data-ttu-id="9f271-184">値</span><span class="sxs-lookup"><span data-stu-id="9f271-184">Value</span></span>|
|---|---|
|[<span data-ttu-id="9f271-185">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="9f271-185">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9f271-186">1.5</span><span class="sxs-lookup"><span data-stu-id="9f271-186">1.5</span></span> |
|[<span data-ttu-id="9f271-187">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="9f271-187">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9f271-188">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9f271-188">ReadItem</span></span>|
|[<span data-ttu-id="9f271-189">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="9f271-189">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9f271-190">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="9f271-190">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="9f271-191">メソッド</span><span class="sxs-lookup"><span data-stu-id="9f271-191">Methods</span></span>

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="9f271-192">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="9f271-192">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="9f271-193">サポートされているイベントのイベント ハンドラーを追加します。</span><span class="sxs-lookup"><span data-stu-id="9f271-193">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="9f271-194">現在、サポートされているイベントの種類は `Office.EventType.ItemChanged` だけです。</span><span class="sxs-lookup"><span data-stu-id="9f271-194">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9f271-195">パラメーター</span><span class="sxs-lookup"><span data-stu-id="9f271-195">Parameters</span></span>

| <span data-ttu-id="9f271-196">名前</span><span class="sxs-lookup"><span data-stu-id="9f271-196">Name</span></span> | <span data-ttu-id="9f271-197">種類</span><span class="sxs-lookup"><span data-stu-id="9f271-197">Type</span></span> | <span data-ttu-id="9f271-198">属性</span><span class="sxs-lookup"><span data-stu-id="9f271-198">Attributes</span></span> | <span data-ttu-id="9f271-199">説明</span><span class="sxs-lookup"><span data-stu-id="9f271-199">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="9f271-200">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="9f271-200">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="9f271-201">ハンドラーを呼び出す必要のあるイベント。</span><span class="sxs-lookup"><span data-stu-id="9f271-201">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="9f271-202">Function</span><span class="sxs-lookup"><span data-stu-id="9f271-202">Function</span></span> || <span data-ttu-id="9f271-p105">イベントを処理する関数。関数は、オブジェクト リテラルである単一パラメーターを受け入れる必要があります。パラメーターの `type` プロパティは、`addHandlerAsync` に渡される `eventType` パラメーターと一致します。</span><span class="sxs-lookup"><span data-stu-id="9f271-p105">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="9f271-206">Object</span><span class="sxs-lookup"><span data-stu-id="9f271-206">Object</span></span> | <span data-ttu-id="9f271-207">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="9f271-207">&lt;optional&gt;</span></span> | <span data-ttu-id="9f271-208">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="9f271-208">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="9f271-209">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="9f271-209">Object</span></span> | <span data-ttu-id="9f271-210">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="9f271-210">&lt;optional&gt;</span></span> | <span data-ttu-id="9f271-211">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="9f271-211">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="9f271-212">function</span><span class="sxs-lookup"><span data-stu-id="9f271-212">function</span></span>| <span data-ttu-id="9f271-213">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="9f271-213">&lt;optional&gt;</span></span>|<span data-ttu-id="9f271-214">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="9f271-214">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9f271-215">要件</span><span class="sxs-lookup"><span data-stu-id="9f271-215">Requirements</span></span>

|<span data-ttu-id="9f271-216">要件</span><span class="sxs-lookup"><span data-stu-id="9f271-216">Requirement</span></span>| <span data-ttu-id="9f271-217">値</span><span class="sxs-lookup"><span data-stu-id="9f271-217">Value</span></span>|
|---|---|
|[<span data-ttu-id="9f271-218">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="9f271-218">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9f271-219">1.5</span><span class="sxs-lookup"><span data-stu-id="9f271-219">1.5</span></span> |
|[<span data-ttu-id="9f271-220">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="9f271-220">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9f271-221">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9f271-221">ReadItem</span></span> |
|[<span data-ttu-id="9f271-222">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="9f271-222">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9f271-223">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="9f271-223">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9f271-224">例</span><span class="sxs-lookup"><span data-stu-id="9f271-224">Example</span></span>

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

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="9f271-225">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="9f271-225">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="9f271-226">REST 形式のアイテム ID を EWS 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="9f271-226">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="9f271-227">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="9f271-227">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="9f271-p106">REST API ([Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) や [Microsoft Graph](https://graph.microsoft.io/) など) で取得されたアイテム ID は、Exchange Web サービス (EWS) に使用される形式とは異なる形式を使用します。`convertToEwsId` メソッドは、REST 形式の ID を EWS 用の適切な形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="9f271-p106">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9f271-230">パラメーター</span><span class="sxs-lookup"><span data-stu-id="9f271-230">Parameters</span></span>

|<span data-ttu-id="9f271-231">名前</span><span class="sxs-lookup"><span data-stu-id="9f271-231">Name</span></span>| <span data-ttu-id="9f271-232">型</span><span class="sxs-lookup"><span data-stu-id="9f271-232">Type</span></span>| <span data-ttu-id="9f271-233">説明</span><span class="sxs-lookup"><span data-stu-id="9f271-233">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="9f271-234">String</span><span class="sxs-lookup"><span data-stu-id="9f271-234">String</span></span>|<span data-ttu-id="9f271-235">Outlook REST API 形式のアイテム ID</span><span class="sxs-lookup"><span data-stu-id="9f271-235">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="9f271-236">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="9f271-236">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.7)|<span data-ttu-id="9f271-237">アイテム ID の取得に使用された Outlook REST API のバージョンを示す値。</span><span class="sxs-lookup"><span data-stu-id="9f271-237">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9f271-238">要件</span><span class="sxs-lookup"><span data-stu-id="9f271-238">Requirements</span></span>

|<span data-ttu-id="9f271-239">要件</span><span class="sxs-lookup"><span data-stu-id="9f271-239">Requirement</span></span>| <span data-ttu-id="9f271-240">値</span><span class="sxs-lookup"><span data-stu-id="9f271-240">Value</span></span>|
|---|---|
|[<span data-ttu-id="9f271-241">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="9f271-241">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9f271-242">1.3</span><span class="sxs-lookup"><span data-stu-id="9f271-242">1.3</span></span>|
|[<span data-ttu-id="9f271-243">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="9f271-243">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9f271-244">制限あり</span><span class="sxs-lookup"><span data-stu-id="9f271-244">Restricted</span></span>|
|[<span data-ttu-id="9f271-245">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="9f271-245">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9f271-246">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="9f271-246">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9f271-247">戻り値:</span><span class="sxs-lookup"><span data-stu-id="9f271-247">Returns:</span></span>

<span data-ttu-id="9f271-248">型:String</span><span class="sxs-lookup"><span data-stu-id="9f271-248">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="9f271-249">例</span><span class="sxs-lookup"><span data-stu-id="9f271-249">Example</span></span>

```js
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-17"></a><span data-ttu-id="9f271-250">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.7)}</span><span class="sxs-lookup"><span data-stu-id="9f271-250">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.7)}</span></span>

<span data-ttu-id="9f271-251">クライアントのローカル時間で時間情報が含まれている辞書を取得します。</span><span class="sxs-lookup"><span data-stu-id="9f271-251">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="9f271-252">デスクトップまたは web 上の Outlook 用メールアプリは、日付と時刻に異なるタイムゾーンを使用できます。</span><span class="sxs-lookup"><span data-stu-id="9f271-252">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times.</span></span> <span data-ttu-id="9f271-253">デスクトップ上の Outlook では、クライアントコンピューターのタイムゾーンが使用されます。Outlook on the web では、Exchange 管理センター (EAC) で設定されているタイムゾーンが使用されます。</span><span class="sxs-lookup"><span data-stu-id="9f271-253">Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC).</span></span> <span data-ttu-id="9f271-254">日付と時刻の値を処理して、ユーザーインターフェイスに表示される値が、ユーザーが期待するタイムゾーンに常に一致するようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="9f271-254">You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="9f271-255">デスクトップクライアント上の Outlook でメールアプリが実行されている`convertToLocalClientTime`場合、このメソッドは、クライアントコンピューターのタイムゾーンに設定された値を持つ dictionary オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="9f271-255">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone.</span></span> <span data-ttu-id="9f271-256">メールアプリが web 上の Outlook で実行されている`convertToLocalClientTime`場合、このメソッドは、EAC で指定されたタイムゾーンに設定された値を持つ dictionary オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="9f271-256">If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9f271-257">パラメーター</span><span class="sxs-lookup"><span data-stu-id="9f271-257">Parameters</span></span>

|<span data-ttu-id="9f271-258">名前</span><span class="sxs-lookup"><span data-stu-id="9f271-258">Name</span></span>| <span data-ttu-id="9f271-259">型</span><span class="sxs-lookup"><span data-stu-id="9f271-259">Type</span></span>| <span data-ttu-id="9f271-260">説明</span><span class="sxs-lookup"><span data-stu-id="9f271-260">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="9f271-261">日付</span><span class="sxs-lookup"><span data-stu-id="9f271-261">Date</span></span>|<span data-ttu-id="9f271-262">日付オブジェクト</span><span class="sxs-lookup"><span data-stu-id="9f271-262">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9f271-263">要件</span><span class="sxs-lookup"><span data-stu-id="9f271-263">Requirements</span></span>

|<span data-ttu-id="9f271-264">要件</span><span class="sxs-lookup"><span data-stu-id="9f271-264">Requirement</span></span>| <span data-ttu-id="9f271-265">値</span><span class="sxs-lookup"><span data-stu-id="9f271-265">Value</span></span>|
|---|---|
|[<span data-ttu-id="9f271-266">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="9f271-266">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9f271-267">1.0</span><span class="sxs-lookup"><span data-stu-id="9f271-267">1.0</span></span>|
|[<span data-ttu-id="9f271-268">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="9f271-268">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9f271-269">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9f271-269">ReadItem</span></span>|
|[<span data-ttu-id="9f271-270">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="9f271-270">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9f271-271">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="9f271-271">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9f271-272">戻り値:</span><span class="sxs-lookup"><span data-stu-id="9f271-272">Returns:</span></span>

<span data-ttu-id="9f271-273">型:[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="9f271-273">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.7)</span></span>

<br>

---
---

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="9f271-274">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="9f271-274">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="9f271-275">EWS 形式のアイテム ID を REST 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="9f271-275">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="9f271-276">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="9f271-276">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="9f271-p109">EWS または `itemId` プロパティで取得されるアイテム ID は、REST API ([Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) や [Microsoft Graph](https://graph.microsoft.io/) など) に使用される形式とは異なる形式を使用します。`convertToRestId` メソッドは、EWS 形式の ID を REST 用の適切な形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="9f271-p109">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9f271-279">パラメーター</span><span class="sxs-lookup"><span data-stu-id="9f271-279">Parameters</span></span>

|<span data-ttu-id="9f271-280">名前</span><span class="sxs-lookup"><span data-stu-id="9f271-280">Name</span></span>| <span data-ttu-id="9f271-281">型</span><span class="sxs-lookup"><span data-stu-id="9f271-281">Type</span></span>| <span data-ttu-id="9f271-282">説明</span><span class="sxs-lookup"><span data-stu-id="9f271-282">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="9f271-283">String</span><span class="sxs-lookup"><span data-stu-id="9f271-283">String</span></span>|<span data-ttu-id="9f271-284">Exchange Web サービス (EWS) 形式のアイテム ID</span><span class="sxs-lookup"><span data-stu-id="9f271-284">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="9f271-285">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="9f271-285">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.7)|<span data-ttu-id="9f271-286">変換後の ID を使用する Outlook REST API のバージョンを示す値。</span><span class="sxs-lookup"><span data-stu-id="9f271-286">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9f271-287">要件</span><span class="sxs-lookup"><span data-stu-id="9f271-287">Requirements</span></span>

|<span data-ttu-id="9f271-288">要件</span><span class="sxs-lookup"><span data-stu-id="9f271-288">Requirement</span></span>| <span data-ttu-id="9f271-289">値</span><span class="sxs-lookup"><span data-stu-id="9f271-289">Value</span></span>|
|---|---|
|[<span data-ttu-id="9f271-290">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="9f271-290">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9f271-291">1.3</span><span class="sxs-lookup"><span data-stu-id="9f271-291">1.3</span></span>|
|[<span data-ttu-id="9f271-292">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="9f271-292">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9f271-293">制限あり</span><span class="sxs-lookup"><span data-stu-id="9f271-293">Restricted</span></span>|
|[<span data-ttu-id="9f271-294">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="9f271-294">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9f271-295">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="9f271-295">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9f271-296">戻り値:</span><span class="sxs-lookup"><span data-stu-id="9f271-296">Returns:</span></span>

<span data-ttu-id="9f271-297">型:String</span><span class="sxs-lookup"><span data-stu-id="9f271-297">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="9f271-298">例</span><span class="sxs-lookup"><span data-stu-id="9f271-298">Example</span></span>

```js
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="9f271-299">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="9f271-299">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="9f271-300">時間情報が含まれているディクショナリから日付オブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="9f271-300">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="9f271-301">`convertToUtcClientTime` メソッドは、ローカルの日付と時刻を含むディクショナリを、ローカルの日付と時刻の正しい値を持つ日付オブジェクトに変換します。</span><span class="sxs-lookup"><span data-stu-id="9f271-301">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9f271-302">パラメーター</span><span class="sxs-lookup"><span data-stu-id="9f271-302">Parameters</span></span>

|<span data-ttu-id="9f271-303">名前</span><span class="sxs-lookup"><span data-stu-id="9f271-303">Name</span></span>| <span data-ttu-id="9f271-304">型</span><span class="sxs-lookup"><span data-stu-id="9f271-304">Type</span></span>| <span data-ttu-id="9f271-305">説明</span><span class="sxs-lookup"><span data-stu-id="9f271-305">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="9f271-306">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="9f271-306">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.7)|<span data-ttu-id="9f271-307">変換するローカル時刻の値。</span><span class="sxs-lookup"><span data-stu-id="9f271-307">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9f271-308">要件</span><span class="sxs-lookup"><span data-stu-id="9f271-308">Requirements</span></span>

|<span data-ttu-id="9f271-309">要件</span><span class="sxs-lookup"><span data-stu-id="9f271-309">Requirement</span></span>| <span data-ttu-id="9f271-310">値</span><span class="sxs-lookup"><span data-stu-id="9f271-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="9f271-311">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="9f271-311">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9f271-312">1.0</span><span class="sxs-lookup"><span data-stu-id="9f271-312">1.0</span></span>|
|[<span data-ttu-id="9f271-313">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="9f271-313">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9f271-314">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9f271-314">ReadItem</span></span>|
|[<span data-ttu-id="9f271-315">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="9f271-315">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9f271-316">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="9f271-316">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9f271-317">戻り値:</span><span class="sxs-lookup"><span data-stu-id="9f271-317">Returns:</span></span>

<span data-ttu-id="9f271-318">時間が UTC で表現された日付オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="9f271-318">A Date object with the time expressed in UTC.</span></span>

<span data-ttu-id="9f271-319">型: Date</span><span class="sxs-lookup"><span data-stu-id="9f271-319">Type: Date</span></span>

##### <a name="example"></a><span data-ttu-id="9f271-320">例</span><span class="sxs-lookup"><span data-stu-id="9f271-320">Example</span></span>

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

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="9f271-321">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="9f271-321">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="9f271-322">既存の予定を表示します。</span><span class="sxs-lookup"><span data-stu-id="9f271-322">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="9f271-323">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="9f271-323">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="9f271-324">`displayAppointmentForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存の予定を開きます。</span><span class="sxs-lookup"><span data-stu-id="9f271-324">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="9f271-325">Outlook on the Mac では、このメソッドを使用して、定期的なアイテムの一部ではない単一の予定を表示したり、定期的なアイテムのマスター予定を表示したりすることはできませんが、一連のインスタンスを表示することはできません。</span><span class="sxs-lookup"><span data-stu-id="9f271-325">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series.</span></span> <span data-ttu-id="9f271-326">これは、Mac 上の Outlook では、定期的なアイテムのインスタンスのプロパティ (アイテム ID を含む) にアクセスできないためです。</span><span class="sxs-lookup"><span data-stu-id="9f271-326">This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="9f271-327">Web 上の Outlook では、フォームの本文が 32 KB 以下の文字である場合にのみ、このメソッドは指定されたフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="9f271-327">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="9f271-328">指定のアイテム識別子が既存の予定を表していない場合は、クライアント コンピューターまたはデバイスで空のウィンドウが開き、エラー メッセージは返されません。</span><span class="sxs-lookup"><span data-stu-id="9f271-328">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9f271-329">パラメーター</span><span class="sxs-lookup"><span data-stu-id="9f271-329">Parameters</span></span>

|<span data-ttu-id="9f271-330">名前</span><span class="sxs-lookup"><span data-stu-id="9f271-330">Name</span></span>| <span data-ttu-id="9f271-331">型</span><span class="sxs-lookup"><span data-stu-id="9f271-331">Type</span></span>| <span data-ttu-id="9f271-332">説明</span><span class="sxs-lookup"><span data-stu-id="9f271-332">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="9f271-333">String</span><span class="sxs-lookup"><span data-stu-id="9f271-333">String</span></span>|<span data-ttu-id="9f271-334">既存の予定の Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="9f271-334">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9f271-335">要件</span><span class="sxs-lookup"><span data-stu-id="9f271-335">Requirements</span></span>

|<span data-ttu-id="9f271-336">要件</span><span class="sxs-lookup"><span data-stu-id="9f271-336">Requirement</span></span>| <span data-ttu-id="9f271-337">値</span><span class="sxs-lookup"><span data-stu-id="9f271-337">Value</span></span>|
|---|---|
|[<span data-ttu-id="9f271-338">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="9f271-338">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9f271-339">1.0</span><span class="sxs-lookup"><span data-stu-id="9f271-339">1.0</span></span>|
|[<span data-ttu-id="9f271-340">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="9f271-340">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9f271-341">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9f271-341">ReadItem</span></span>|
|[<span data-ttu-id="9f271-342">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="9f271-342">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9f271-343">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="9f271-343">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9f271-344">例</span><span class="sxs-lookup"><span data-stu-id="9f271-344">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

<br>

---
---

#### <a name="displaymessageformitemid"></a><span data-ttu-id="9f271-345">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="9f271-345">displayMessageForm(itemId)</span></span>

<span data-ttu-id="9f271-346">既存のメッセージを表示します。</span><span class="sxs-lookup"><span data-stu-id="9f271-346">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="9f271-347">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="9f271-347">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="9f271-348">`displayMessageForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存のメッセージを開きます。</span><span class="sxs-lookup"><span data-stu-id="9f271-348">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="9f271-349">Web 上の Outlook では、フォームの本文が 32 KB の文字数以下の場合にのみ、このメソッドは指定されたフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="9f271-349">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="9f271-350">指定のアイテム識別子が既存のメッセージを表していない場合は、クライアント コンピューターにはメッセージは表示されず、エラー メッセージも返されません。</span><span class="sxs-lookup"><span data-stu-id="9f271-350">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="9f271-p111">予定を表す `itemId` を含む `displayMessageForm` を使用しないでください。既存の予定を表示するには、`displayAppointmentForm` メソッドを使用します。新しい予定を作成するフォームを表示するには、`displayNewAppointmentForm` メソッドを使用します。</span><span class="sxs-lookup"><span data-stu-id="9f271-p111">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9f271-353">パラメーター</span><span class="sxs-lookup"><span data-stu-id="9f271-353">Parameters</span></span>

|<span data-ttu-id="9f271-354">名前</span><span class="sxs-lookup"><span data-stu-id="9f271-354">Name</span></span>| <span data-ttu-id="9f271-355">型</span><span class="sxs-lookup"><span data-stu-id="9f271-355">Type</span></span>| <span data-ttu-id="9f271-356">説明</span><span class="sxs-lookup"><span data-stu-id="9f271-356">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="9f271-357">String</span><span class="sxs-lookup"><span data-stu-id="9f271-357">String</span></span>|<span data-ttu-id="9f271-358">既存のメッセージの Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="9f271-358">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9f271-359">要件</span><span class="sxs-lookup"><span data-stu-id="9f271-359">Requirements</span></span>

|<span data-ttu-id="9f271-360">要件</span><span class="sxs-lookup"><span data-stu-id="9f271-360">Requirement</span></span>| <span data-ttu-id="9f271-361">値</span><span class="sxs-lookup"><span data-stu-id="9f271-361">Value</span></span>|
|---|---|
|[<span data-ttu-id="9f271-362">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="9f271-362">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9f271-363">1.0</span><span class="sxs-lookup"><span data-stu-id="9f271-363">1.0</span></span>|
|[<span data-ttu-id="9f271-364">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="9f271-364">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9f271-365">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9f271-365">ReadItem</span></span>|
|[<span data-ttu-id="9f271-366">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="9f271-366">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9f271-367">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="9f271-367">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9f271-368">例</span><span class="sxs-lookup"><span data-stu-id="9f271-368">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

<br>

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="9f271-369">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="9f271-369">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="9f271-370">新しい予定を作成するためのフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="9f271-370">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="9f271-371">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="9f271-371">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="9f271-p112">`displayNewAppointmentForm` メソッドを使用すると、ユーザーが新しい予定または会議を作成できるフォームが開きます。パラメーターを指定すると、予定のフォーム フィールドにパラメーターの内容が自動的に設定されます。</span><span class="sxs-lookup"><span data-stu-id="9f271-p112">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="9f271-374">Outlook on the web およびモバイルデバイスでは、このメソッドは常に出席者フィールドを含むフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="9f271-374">In Outlook on the web and mobile devices, this method always displays a form with an attendees field.</span></span> <span data-ttu-id="9f271-375">入力引数として出席者を指定しないと、このメソッドにより **[保存]** ボタンのあるフォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="9f271-375">If you do not specify any attendees as input arguments, the method displays a form with a **Save** button.</span></span> <span data-ttu-id="9f271-376">出席者を指定した場合には、フォームにその出席者と **[送信]** ボタンが表示されます。</span><span class="sxs-lookup"><span data-stu-id="9f271-376">If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="9f271-p114">Outlook リッチ クライアントと Outlook RT で、`requiredAttendees`、`optionalAttendees`、または `resources` パラメーターに出席者またはリソースを指定し、このメソッドを実行すると、**[送信]** ボタンがある会議フォームが表示されます。受信者を指定せずにこのメソッドを実行すると、**[保存して閉じる]** ボタンがある予定フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="9f271-p114">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="9f271-379">パラメーターのいずれかが指定のサイズ制限を超える場合、または不明なパラメーター名が指定されている場合は、例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="9f271-379">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9f271-380">パラメーター</span><span class="sxs-lookup"><span data-stu-id="9f271-380">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="9f271-381">すべてのパラメーターは省略可能です。</span><span class="sxs-lookup"><span data-stu-id="9f271-381">All parameters are optional.</span></span>

|<span data-ttu-id="9f271-382">名前</span><span class="sxs-lookup"><span data-stu-id="9f271-382">Name</span></span>| <span data-ttu-id="9f271-383">型</span><span class="sxs-lookup"><span data-stu-id="9f271-383">Type</span></span>| <span data-ttu-id="9f271-384">説明</span><span class="sxs-lookup"><span data-stu-id="9f271-384">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="9f271-385">Object</span><span class="sxs-lookup"><span data-stu-id="9f271-385">Object</span></span> | <span data-ttu-id="9f271-386">新しい予定を記述するパラメーターのディクショナリ。</span><span class="sxs-lookup"><span data-stu-id="9f271-386">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="9f271-387">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span><span class="sxs-lookup"><span data-stu-id="9f271-387">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span></span> | <span data-ttu-id="9f271-p115">予定に必要な各出席者について、メール アドレスを含む文字列の配列、または `EmailAddressDetails` オブジェクトを含む配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="9f271-p115">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="9f271-390">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span><span class="sxs-lookup"><span data-stu-id="9f271-390">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span></span> | <span data-ttu-id="9f271-p116">予定の各任意出席者について、メール アドレスを含む文字列の配列、または `EmailAddressDetails` オブジェクトを含む配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="9f271-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="9f271-393">日付</span><span class="sxs-lookup"><span data-stu-id="9f271-393">Date</span></span> | <span data-ttu-id="9f271-394">予定の開始日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="9f271-394">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="9f271-395">日付</span><span class="sxs-lookup"><span data-stu-id="9f271-395">Date</span></span> | <span data-ttu-id="9f271-396">予定の終了日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="9f271-396">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="9f271-397">String</span><span class="sxs-lookup"><span data-stu-id="9f271-397">String</span></span> | <span data-ttu-id="9f271-p117">予定の場所を含む文字列。文字列は最大 255 文字に制限されます。</span><span class="sxs-lookup"><span data-stu-id="9f271-p117">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="9f271-400">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="9f271-400">Array.&lt;String&gt;</span></span> | <span data-ttu-id="9f271-p118">予定に必要なリソースを含む文字列の配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="9f271-p118">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="9f271-403">String</span><span class="sxs-lookup"><span data-stu-id="9f271-403">String</span></span> | <span data-ttu-id="9f271-p119">予定の件名を含む文字列です。文字列は最大 255 文字に制限されます。</span><span class="sxs-lookup"><span data-stu-id="9f271-p119">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="9f271-406">String</span><span class="sxs-lookup"><span data-stu-id="9f271-406">String</span></span> | <span data-ttu-id="9f271-p120">予定の本文。本文の内容は、最大サイズが 32 KB に制限されます。</span><span class="sxs-lookup"><span data-stu-id="9f271-p120">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="9f271-409">要件</span><span class="sxs-lookup"><span data-stu-id="9f271-409">Requirements</span></span>

|<span data-ttu-id="9f271-410">要件</span><span class="sxs-lookup"><span data-stu-id="9f271-410">Requirement</span></span>| <span data-ttu-id="9f271-411">値</span><span class="sxs-lookup"><span data-stu-id="9f271-411">Value</span></span>|
|---|---|
|[<span data-ttu-id="9f271-412">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="9f271-412">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9f271-413">1.0</span><span class="sxs-lookup"><span data-stu-id="9f271-413">1.0</span></span>|
|[<span data-ttu-id="9f271-414">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="9f271-414">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9f271-415">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9f271-415">ReadItem</span></span>|
|[<span data-ttu-id="9f271-416">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="9f271-416">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9f271-417">読み取り</span><span class="sxs-lookup"><span data-stu-id="9f271-417">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9f271-418">例</span><span class="sxs-lookup"><span data-stu-id="9f271-418">Example</span></span>

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

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="9f271-419">displayNewMessageForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="9f271-419">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="9f271-420">新しいメッセージを作成するためのフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="9f271-420">Displays a form for creating a new message.</span></span>

<span data-ttu-id="9f271-p121">The `displayNewMessageForm` method opens a form that enables the user to create a new message. If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span><span class="sxs-lookup"><span data-stu-id="9f271-p121">The `displayNewMessageForm` method opens a form that enables the user to create a new message. If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="9f271-423">パラメーターのいずれかが指定のサイズ制限を超える場合、または不明なパラメーター名が指定されている場合は、例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="9f271-423">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9f271-424">パラメーター</span><span class="sxs-lookup"><span data-stu-id="9f271-424">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="9f271-425">すべてのパラメーターは省略可能です。</span><span class="sxs-lookup"><span data-stu-id="9f271-425">All parameters are optional.</span></span>

|<span data-ttu-id="9f271-426">名前</span><span class="sxs-lookup"><span data-stu-id="9f271-426">Name</span></span>| <span data-ttu-id="9f271-427">型</span><span class="sxs-lookup"><span data-stu-id="9f271-427">Type</span></span>| <span data-ttu-id="9f271-428">説明</span><span class="sxs-lookup"><span data-stu-id="9f271-428">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="9f271-429">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="9f271-429">Object</span></span> | <span data-ttu-id="9f271-430">新しいメッセージを記述するパラメータの辞書。</span><span class="sxs-lookup"><span data-stu-id="9f271-430">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="9f271-431">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span><span class="sxs-lookup"><span data-stu-id="9f271-431">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span></span> | <span data-ttu-id="9f271-p122">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line. The array is limited to a maximum of 100 entries.</span><span class="sxs-lookup"><span data-stu-id="9f271-p122">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="9f271-434">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span><span class="sxs-lookup"><span data-stu-id="9f271-434">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span></span> | <span data-ttu-id="9f271-p123">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line. The array is limited to a maximum of 100 entries.</span><span class="sxs-lookup"><span data-stu-id="9f271-p123">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="9f271-437">配列。&lt;文字列&gt; | 配列。&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span><span class="sxs-lookup"><span data-stu-id="9f271-437">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span></span> | <span data-ttu-id="9f271-p124">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line. The array is limited to a maximum of 100 entries.</span><span class="sxs-lookup"><span data-stu-id="9f271-p124">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="9f271-440">String</span><span class="sxs-lookup"><span data-stu-id="9f271-440">String</span></span> | <span data-ttu-id="9f271-p125">A string containing the subject of the message. The string is limited to a maximum of 255 characters.</span><span class="sxs-lookup"><span data-stu-id="9f271-p125">A string containing the subject of the message. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="9f271-443">String</span><span class="sxs-lookup"><span data-stu-id="9f271-443">String</span></span> | <span data-ttu-id="9f271-p126">The HTML body of the message. The body content is limited to a maximum size of 32 KB.</span><span class="sxs-lookup"><span data-stu-id="9f271-p126">The HTML body of the message. The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="9f271-446">配列。&lt;オブジェクト&gt;</span><span class="sxs-lookup"><span data-stu-id="9f271-446">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="9f271-447">添付ファイルまたは添付アイテムである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="9f271-447">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="9f271-448">String</span><span class="sxs-lookup"><span data-stu-id="9f271-448">String</span></span> | <span data-ttu-id="9f271-p127">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="9f271-p127">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="9f271-451">String</span><span class="sxs-lookup"><span data-stu-id="9f271-451">String</span></span> | <span data-ttu-id="9f271-452">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="9f271-452">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="9f271-453">文字列</span><span class="sxs-lookup"><span data-stu-id="9f271-453">String</span></span> | <span data-ttu-id="9f271-p128">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="9f271-p128">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="9f271-456">ブール値</span><span class="sxs-lookup"><span data-stu-id="9f271-456">Boolean</span></span> | <span data-ttu-id="9f271-p129">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="9f271-p129">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="9f271-459">String</span><span class="sxs-lookup"><span data-stu-id="9f271-459">String</span></span> | <span data-ttu-id="9f271-p130">Only used if `type` is set to `item`. The EWS item id of the existing e-mail you want to attach to the new message. This is a string up to 100 characters.</span><span class="sxs-lookup"><span data-stu-id="9f271-p130">Only used if `type` is set to `item`. The EWS item id of the existing e-mail you want to attach to the new message. This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="9f271-463">要件</span><span class="sxs-lookup"><span data-stu-id="9f271-463">Requirements</span></span>

|<span data-ttu-id="9f271-464">要件</span><span class="sxs-lookup"><span data-stu-id="9f271-464">Requirement</span></span>| <span data-ttu-id="9f271-465">値</span><span class="sxs-lookup"><span data-stu-id="9f271-465">Value</span></span>|
|---|---|
|[<span data-ttu-id="9f271-466">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="9f271-466">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9f271-467">1.6</span><span class="sxs-lookup"><span data-stu-id="9f271-467">1.6</span></span> |
|[<span data-ttu-id="9f271-468">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="9f271-468">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9f271-469">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9f271-469">ReadItem</span></span>|
|[<span data-ttu-id="9f271-470">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="9f271-470">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9f271-471">読み取り</span><span class="sxs-lookup"><span data-stu-id="9f271-471">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9f271-472">例</span><span class="sxs-lookup"><span data-stu-id="9f271-472">Example</span></span>

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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="9f271-473">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="9f271-473">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="9f271-474">REST API または Exchange Web サービスを呼び出すために使用するトークンを含む文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="9f271-474">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="9f271-p131">`getCallbackTokenAsync` メソッドは、ユーザーのメールボックスをホストする Exchange Server から不透明なトークンを取得する非同期の呼び出しを行います。コールバック トークンの有効期間は 5 分です。</span><span class="sxs-lookup"><span data-stu-id="9f271-p131">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="9f271-477">可能であれば、アドインでは Exchange Web サービスの代わりに REST API を使用することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="9f271-477">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span>

<span data-ttu-id="9f271-478">**REST トークン**</span><span class="sxs-lookup"><span data-stu-id="9f271-478">**REST Tokens**</span></span>

<span data-ttu-id="9f271-p132">REST トークンが要求された場合 (`options.isRest = true`)、結果トークンは Exchange Web サービスの呼び出しを認証するためには機能しません。アドインがマニフェストで [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) アクセス許可を指定していない限り、トークンの範囲は現在のアイテムとその添付ファイルへの読み取り専用アクセスに制限されます。`ReadWriteMailbox` アクセス許可が指定されている場合は、結果トークンは、メールを送信する機能など、メール、カレンダー、連絡先への読み取り/書き込みアクセスを付与します。</span><span class="sxs-lookup"><span data-stu-id="9f271-p132">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="9f271-482">アドインでは、`restUrl` プロパティを使用して、REST API 呼び出しを行うときに使用する正しい URL を決定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="9f271-482">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="9f271-483">**EWS トークン**</span><span class="sxs-lookup"><span data-stu-id="9f271-483">**EWS Tokens**</span></span>

<span data-ttu-id="9f271-p133">EWS トークンが要求された場合 (`options.isRest = false`)、結果トークンは REST API 呼び出しを認証するためには機能しません。トークンの範囲は、現在のアイテムへのアクセスに制限されます。</span><span class="sxs-lookup"><span data-stu-id="9f271-p133">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="9f271-486">アドインでは、`ewsUrl` プロパティを使用して、EWS 呼び出しを行うときに使用する正しい URL を決定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="9f271-486">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9f271-487">パラメーター</span><span class="sxs-lookup"><span data-stu-id="9f271-487">Parameters</span></span>

|<span data-ttu-id="9f271-488">名前</span><span class="sxs-lookup"><span data-stu-id="9f271-488">Name</span></span>| <span data-ttu-id="9f271-489">型</span><span class="sxs-lookup"><span data-stu-id="9f271-489">Type</span></span>| <span data-ttu-id="9f271-490">属性</span><span class="sxs-lookup"><span data-stu-id="9f271-490">Attributes</span></span>| <span data-ttu-id="9f271-491">説明</span><span class="sxs-lookup"><span data-stu-id="9f271-491">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="9f271-492">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="9f271-492">Object</span></span> | <span data-ttu-id="9f271-493">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="9f271-493">&lt;optional&gt;</span></span> | <span data-ttu-id="9f271-494">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="9f271-494">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="9f271-495">Boolean</span><span class="sxs-lookup"><span data-stu-id="9f271-495">Boolean</span></span> |  <span data-ttu-id="9f271-496">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="9f271-496">&lt;optional&gt;</span></span> | <span data-ttu-id="9f271-p134">提供されたトークンを Outlook REST API または Exchange Web サービスに使用するかどうかを決定します。既定値は、`false` です。</span><span class="sxs-lookup"><span data-stu-id="9f271-p134">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="9f271-499">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="9f271-499">Object</span></span> |  <span data-ttu-id="9f271-500">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="9f271-500">&lt;optional&gt;</span></span> | <span data-ttu-id="9f271-501">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="9f271-501">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="9f271-502">function</span><span class="sxs-lookup"><span data-stu-id="9f271-502">function</span></span>||<span data-ttu-id="9f271-503">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="9f271-503">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="9f271-504">トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="9f271-504">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="9f271-505">エラーが発生した場合は`asyncResult.error` 、 `asyncResult.diagnostics`プロパティとプロパティによって追加情報が提供されることがあります。</span><span class="sxs-lookup"><span data-stu-id="9f271-505">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|

##### <a name="errors"></a><span data-ttu-id="9f271-506">エラー</span><span class="sxs-lookup"><span data-stu-id="9f271-506">Errors</span></span>

|<span data-ttu-id="9f271-507">エラー コード</span><span class="sxs-lookup"><span data-stu-id="9f271-507">Error code</span></span>|<span data-ttu-id="9f271-508">説明</span><span class="sxs-lookup"><span data-stu-id="9f271-508">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="9f271-509">要求が失敗しました。</span><span class="sxs-lookup"><span data-stu-id="9f271-509">The request has failed.</span></span> <span data-ttu-id="9f271-510">HTTP エラーコードについては、diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="9f271-510">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="9f271-511">Exchange サーバーがエラーを返しました。</span><span class="sxs-lookup"><span data-stu-id="9f271-511">The Exchange server returned an error.</span></span> <span data-ttu-id="9f271-512">詳細については、「diagnostics オブジェクト」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="9f271-512">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="9f271-513">ユーザーがネットワークに接続されていません。</span><span class="sxs-lookup"><span data-stu-id="9f271-513">The user is no longer connected to the network.</span></span> <span data-ttu-id="9f271-514">ネットワーク接続を確認し、もう一度実行してください。</span><span class="sxs-lookup"><span data-stu-id="9f271-514">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9f271-515">要件</span><span class="sxs-lookup"><span data-stu-id="9f271-515">Requirements</span></span>

|<span data-ttu-id="9f271-516">要件</span><span class="sxs-lookup"><span data-stu-id="9f271-516">Requirement</span></span>| <span data-ttu-id="9f271-517">値</span><span class="sxs-lookup"><span data-stu-id="9f271-517">Value</span></span>|
|---|---|
|[<span data-ttu-id="9f271-518">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="9f271-518">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9f271-519">1.5</span><span class="sxs-lookup"><span data-stu-id="9f271-519">1.5</span></span> |
|[<span data-ttu-id="9f271-520">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="9f271-520">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9f271-521">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9f271-521">ReadItem</span></span>|
|[<span data-ttu-id="9f271-522">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="9f271-522">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9f271-523">新規作成と閲覧</span><span class="sxs-lookup"><span data-stu-id="9f271-523">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="9f271-524">例</span><span class="sxs-lookup"><span data-stu-id="9f271-524">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="9f271-525">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="9f271-525">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="9f271-526">Exchange Server から添付ファイルやアイテムを取得するために使うトークンを含む文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="9f271-526">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="9f271-p138">`getCallbackTokenAsync` メソッドは、ユーザーのメールボックスをホストする Exchange Server から不透明なトークンを取得する非同期の呼び出しを行います。コールバック トークンの有効期間は 5 分です。</span><span class="sxs-lookup"><span data-stu-id="9f271-p138">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="9f271-p139">トークンと予定の識別子またはアイテムの識別子をサードパーティ システムに渡すことができます。サードパーティ システムは、トークンをベアラー承認トークンとして使用し、Exchange Web サービス (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) または [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) 操作を呼び出して、添付ファイルまたはアイテムを返します。たとえば、[選択したアイテムから添付ファイルを取得する](/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="9f271-p139">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="9f271-532">アプリが閲覧モードで `getCallbackTokenAsync` メソッドを呼び出すには、アプリのマニフェスト内に **ReadItem** アクセス許可が指定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="9f271-532">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="9f271-p140">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出してアイテムの識別子を `getCallbackTokenAsync` メソッドに渡す必要があります。アプリには、`saveAsync` メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="9f271-p140">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9f271-535">パラメーター</span><span class="sxs-lookup"><span data-stu-id="9f271-535">Parameters</span></span>

|<span data-ttu-id="9f271-536">名前</span><span class="sxs-lookup"><span data-stu-id="9f271-536">Name</span></span>| <span data-ttu-id="9f271-537">型</span><span class="sxs-lookup"><span data-stu-id="9f271-537">Type</span></span>| <span data-ttu-id="9f271-538">属性</span><span class="sxs-lookup"><span data-stu-id="9f271-538">Attributes</span></span>| <span data-ttu-id="9f271-539">説明</span><span class="sxs-lookup"><span data-stu-id="9f271-539">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="9f271-540">関数</span><span class="sxs-lookup"><span data-stu-id="9f271-540">function</span></span>||<span data-ttu-id="9f271-541">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="9f271-541">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="9f271-542">トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="9f271-542">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="9f271-543">エラーが発生した場合は`asyncResult.error` 、 `asyncResult.diagnostics`プロパティとプロパティによって追加情報が提供されることがあります。</span><span class="sxs-lookup"><span data-stu-id="9f271-543">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="9f271-544">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="9f271-544">Object</span></span>| <span data-ttu-id="9f271-545">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="9f271-545">&lt;optional&gt;</span></span>|<span data-ttu-id="9f271-546">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="9f271-546">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="9f271-547">エラー</span><span class="sxs-lookup"><span data-stu-id="9f271-547">Errors</span></span>

|<span data-ttu-id="9f271-548">エラー コード</span><span class="sxs-lookup"><span data-stu-id="9f271-548">Error code</span></span>|<span data-ttu-id="9f271-549">説明</span><span class="sxs-lookup"><span data-stu-id="9f271-549">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="9f271-550">要求が失敗しました。</span><span class="sxs-lookup"><span data-stu-id="9f271-550">The request has failed.</span></span> <span data-ttu-id="9f271-551">HTTP エラーコードについては、diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="9f271-551">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="9f271-552">Exchange サーバーがエラーを返しました。</span><span class="sxs-lookup"><span data-stu-id="9f271-552">The Exchange server returned an error.</span></span> <span data-ttu-id="9f271-553">詳細については、「diagnostics オブジェクト」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="9f271-553">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="9f271-554">ユーザーがネットワークに接続されていません。</span><span class="sxs-lookup"><span data-stu-id="9f271-554">The user is no longer connected to the network.</span></span> <span data-ttu-id="9f271-555">ネットワーク接続を確認し、もう一度実行してください。</span><span class="sxs-lookup"><span data-stu-id="9f271-555">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9f271-556">要件</span><span class="sxs-lookup"><span data-stu-id="9f271-556">Requirements</span></span>

|<span data-ttu-id="9f271-557">要件</span><span class="sxs-lookup"><span data-stu-id="9f271-557">Requirement</span></span>| <span data-ttu-id="9f271-558">値</span><span class="sxs-lookup"><span data-stu-id="9f271-558">Value</span></span>|
|---|---|
|[<span data-ttu-id="9f271-559">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="9f271-559">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9f271-560">1.0</span><span class="sxs-lookup"><span data-stu-id="9f271-560">1.0</span></span>|
|[<span data-ttu-id="9f271-561">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="9f271-561">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9f271-562">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9f271-562">ReadItem</span></span>|
|[<span data-ttu-id="9f271-563">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="9f271-563">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9f271-564">新規作成と閲覧</span><span class="sxs-lookup"><span data-stu-id="9f271-564">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="9f271-565">例</span><span class="sxs-lookup"><span data-stu-id="9f271-565">Example</span></span>

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

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="9f271-566">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="9f271-566">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="9f271-567">ユーザーと Office アドインを識別するトークンを取得します。</span><span class="sxs-lookup"><span data-stu-id="9f271-567">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="9f271-568">`getUserIdentityTokenAsync` メソッドは、[アドインとユーザーをサード パーティのシステムで識別して認証](/outlook/add-ins/authentication)することのできるトークンを返します。</span><span class="sxs-lookup"><span data-stu-id="9f271-568">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="9f271-569">パラメーター</span><span class="sxs-lookup"><span data-stu-id="9f271-569">Parameters</span></span>

|<span data-ttu-id="9f271-570">名前</span><span class="sxs-lookup"><span data-stu-id="9f271-570">Name</span></span>| <span data-ttu-id="9f271-571">型</span><span class="sxs-lookup"><span data-stu-id="9f271-571">Type</span></span>| <span data-ttu-id="9f271-572">属性</span><span class="sxs-lookup"><span data-stu-id="9f271-572">Attributes</span></span>| <span data-ttu-id="9f271-573">説明</span><span class="sxs-lookup"><span data-stu-id="9f271-573">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="9f271-574">関数</span><span class="sxs-lookup"><span data-stu-id="9f271-574">function</span></span>||<span data-ttu-id="9f271-575">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="9f271-575">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="9f271-576">トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="9f271-576">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="9f271-577">エラーが発生した場合は`asyncResult.error` 、 `asyncResult.diagnostics`プロパティとプロパティによって追加情報が提供されることがあります。</span><span class="sxs-lookup"><span data-stu-id="9f271-577">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="9f271-578">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="9f271-578">Object</span></span>| <span data-ttu-id="9f271-579">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="9f271-579">&lt;optional&gt;</span></span>|<span data-ttu-id="9f271-580">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="9f271-580">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="9f271-581">エラー</span><span class="sxs-lookup"><span data-stu-id="9f271-581">Errors</span></span>

|<span data-ttu-id="9f271-582">エラー コード</span><span class="sxs-lookup"><span data-stu-id="9f271-582">Error code</span></span>|<span data-ttu-id="9f271-583">説明</span><span class="sxs-lookup"><span data-stu-id="9f271-583">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="9f271-584">要求が失敗しました。</span><span class="sxs-lookup"><span data-stu-id="9f271-584">The request has failed.</span></span> <span data-ttu-id="9f271-585">HTTP エラーコードについては、diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="9f271-585">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="9f271-586">Exchange サーバーがエラーを返しました。</span><span class="sxs-lookup"><span data-stu-id="9f271-586">The Exchange server returned an error.</span></span> <span data-ttu-id="9f271-587">詳細については、「diagnostics オブジェクト」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="9f271-587">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="9f271-588">ユーザーがネットワークに接続されていません。</span><span class="sxs-lookup"><span data-stu-id="9f271-588">The user is no longer connected to the network.</span></span> <span data-ttu-id="9f271-589">ネットワーク接続を確認し、もう一度実行してください。</span><span class="sxs-lookup"><span data-stu-id="9f271-589">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9f271-590">要件</span><span class="sxs-lookup"><span data-stu-id="9f271-590">Requirements</span></span>

|<span data-ttu-id="9f271-591">要件</span><span class="sxs-lookup"><span data-stu-id="9f271-591">Requirement</span></span>| <span data-ttu-id="9f271-592">値</span><span class="sxs-lookup"><span data-stu-id="9f271-592">Value</span></span>|
|---|---|
|[<span data-ttu-id="9f271-593">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="9f271-593">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9f271-594">1.0</span><span class="sxs-lookup"><span data-stu-id="9f271-594">1.0</span></span>|
|[<span data-ttu-id="9f271-595">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="9f271-595">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9f271-596">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9f271-596">ReadItem</span></span>|
|[<span data-ttu-id="9f271-597">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="9f271-597">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9f271-598">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="9f271-598">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9f271-599">例</span><span class="sxs-lookup"><span data-stu-id="9f271-599">Example</span></span>

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

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="9f271-600">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="9f271-600">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="9f271-601">ユーザーのメールボックスをホストしている Exchange サーバー上の Exchange Web サービス (EWS) のサービスに対して非同期の要求を行います。</span><span class="sxs-lookup"><span data-stu-id="9f271-601">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="9f271-602">このメソッドは、次のシナリオではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="9f271-602">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="9f271-603">Outlook on iOS または Android</span><span class="sxs-lookup"><span data-stu-id="9f271-603">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="9f271-604">アドインが Gmail のメールボックスに読み込まれる場合</span><span class="sxs-lookup"><span data-stu-id="9f271-604">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="9f271-605">このような場合は、アドインでは [REST API を使用](/outlook/add-ins/use-rest-api)して、代わりにユーザーのメールボックスにアクセスする必要があります。</span><span class="sxs-lookup"><span data-stu-id="9f271-605">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="9f271-p147">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange. See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span><span class="sxs-lookup"><span data-stu-id="9f271-p147">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange. See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="9f271-608">`makeEwsRequestAsync` メソッドでは、フォルダー関連アイテムを要求できません。</span><span class="sxs-lookup"><span data-stu-id="9f271-608">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="9f271-609">XML 要求では UTF-8 エンコードを指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="9f271-609">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="9f271-p148">`makeEwsRequestAsync` メソッドを使用するには、アドインに **ReadWriteMailbox** アクセス許可が必要です。**ReadWriteMailbox** アクセス許可と、`makeEwsRequestAsync` メソッドで呼び出せる EWS 操作の使い方については、「[ユーザーのメールボックスへのメール アドイン アクセスのアクセス許可を指定する](/outlook/add-ins/understanding-outlook-add-in-permissions)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="9f271-p148">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="9f271-612">サーバー管理者は、クライアント アクセス サーバーの EWS ディレクトリで `OAuthAuthentication` を true に設定して、`makeEwsRequestAsync` メソッドで EWS 要求を行うことができるようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="9f271-612">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="9f271-613">バージョンの相違点</span><span class="sxs-lookup"><span data-stu-id="9f271-613">Version differences</span></span>

<span data-ttu-id="9f271-614">バージョン 15.0.4535.1004 より前のバージョンの Outlook で実行しているメール アプリで `makeEwsRequestAsync` メソッドを使う場合は、エンコード値を `ISO-8859-1` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="9f271-614">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="9f271-p149">Outlook on the web でメール アプリを実行している場合は、エンコード値を設定する必要はありません。mailbox.diagnostics.hostName プロパティを使って、メール アプリを Outlook で実行しているのか、Outlook on the web で実行しているのかを確認できます。mailbox.diagnostics.hostVersion プロパティを使って、どのバージョンの Outlook を使って実行しているかを確認できます。</span><span class="sxs-lookup"><span data-stu-id="9f271-p149">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9f271-618">パラメーター</span><span class="sxs-lookup"><span data-stu-id="9f271-618">Parameters</span></span>

|<span data-ttu-id="9f271-619">名前</span><span class="sxs-lookup"><span data-stu-id="9f271-619">Name</span></span>| <span data-ttu-id="9f271-620">型</span><span class="sxs-lookup"><span data-stu-id="9f271-620">Type</span></span>| <span data-ttu-id="9f271-621">属性</span><span class="sxs-lookup"><span data-stu-id="9f271-621">Attributes</span></span>| <span data-ttu-id="9f271-622">説明</span><span class="sxs-lookup"><span data-stu-id="9f271-622">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="9f271-623">String</span><span class="sxs-lookup"><span data-stu-id="9f271-623">String</span></span>||<span data-ttu-id="9f271-624">EWS 要求です。</span><span class="sxs-lookup"><span data-stu-id="9f271-624">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="9f271-625">function</span><span class="sxs-lookup"><span data-stu-id="9f271-625">function</span></span>||<span data-ttu-id="9f271-626">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="9f271-626">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="9f271-p150">The XML result of the EWS call is provided as a string in the `asyncResult.value` property. If the result exceeds 1 MB in size, an error message is returned instead.</span><span class="sxs-lookup"><span data-stu-id="9f271-p150">The XML result of the EWS call is provided as a string in the `asyncResult.value` property. If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="9f271-629">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="9f271-629">Object</span></span>| <span data-ttu-id="9f271-630">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="9f271-630">&lt;optional&gt;</span></span>|<span data-ttu-id="9f271-631">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="9f271-631">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9f271-632">要件</span><span class="sxs-lookup"><span data-stu-id="9f271-632">Requirements</span></span>

|<span data-ttu-id="9f271-633">要件</span><span class="sxs-lookup"><span data-stu-id="9f271-633">Requirement</span></span>| <span data-ttu-id="9f271-634">値</span><span class="sxs-lookup"><span data-stu-id="9f271-634">Value</span></span>|
|---|---|
|[<span data-ttu-id="9f271-635">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="9f271-635">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9f271-636">1.0</span><span class="sxs-lookup"><span data-stu-id="9f271-636">1.0</span></span>|
|[<span data-ttu-id="9f271-637">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="9f271-637">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9f271-638">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="9f271-638">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="9f271-639">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="9f271-639">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9f271-640">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="9f271-640">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9f271-641">例</span><span class="sxs-lookup"><span data-stu-id="9f271-641">Example</span></span>

<span data-ttu-id="9f271-642">次の例は、`GetItem` 操作を使ってアイテムの件名を取得するため、`makeEwsRequestAsync` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="9f271-642">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="9f271-643">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="9f271-643">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="9f271-644">サポートされているイベントの種類のイベント ハンドラーを削除します。</span><span class="sxs-lookup"><span data-stu-id="9f271-644">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="9f271-645">現在、サポートされているイベントの種類は `Office.EventType.ItemChanged` だけです。</span><span class="sxs-lookup"><span data-stu-id="9f271-645">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9f271-646">パラメーター</span><span class="sxs-lookup"><span data-stu-id="9f271-646">Parameters</span></span>

| <span data-ttu-id="9f271-647">名前</span><span class="sxs-lookup"><span data-stu-id="9f271-647">Name</span></span> | <span data-ttu-id="9f271-648">型</span><span class="sxs-lookup"><span data-stu-id="9f271-648">Type</span></span> | <span data-ttu-id="9f271-649">属性</span><span class="sxs-lookup"><span data-stu-id="9f271-649">Attributes</span></span> | <span data-ttu-id="9f271-650">説明</span><span class="sxs-lookup"><span data-stu-id="9f271-650">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="9f271-651">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="9f271-651">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="9f271-652">ハンドラーを取り消すイベント。</span><span class="sxs-lookup"><span data-stu-id="9f271-652">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="9f271-653">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="9f271-653">Object</span></span> | <span data-ttu-id="9f271-654">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="9f271-654">&lt;optional&gt;</span></span> | <span data-ttu-id="9f271-655">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="9f271-655">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="9f271-656">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="9f271-656">Object</span></span> | <span data-ttu-id="9f271-657">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="9f271-657">&lt;optional&gt;</span></span> | <span data-ttu-id="9f271-658">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="9f271-658">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="9f271-659">function</span><span class="sxs-lookup"><span data-stu-id="9f271-659">function</span></span>| <span data-ttu-id="9f271-660">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="9f271-660">&lt;optional&gt;</span></span>|<span data-ttu-id="9f271-661">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="9f271-661">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9f271-662">要件</span><span class="sxs-lookup"><span data-stu-id="9f271-662">Requirements</span></span>

|<span data-ttu-id="9f271-663">要件</span><span class="sxs-lookup"><span data-stu-id="9f271-663">Requirement</span></span>| <span data-ttu-id="9f271-664">値</span><span class="sxs-lookup"><span data-stu-id="9f271-664">Value</span></span>|
|---|---|
|[<span data-ttu-id="9f271-665">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="9f271-665">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9f271-666">1.5</span><span class="sxs-lookup"><span data-stu-id="9f271-666">1.5</span></span> |
|[<span data-ttu-id="9f271-667">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="9f271-667">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9f271-668">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9f271-668">ReadItem</span></span> |
|[<span data-ttu-id="9f271-669">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="9f271-669">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9f271-670">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="9f271-670">Compose or Read</span></span>|
