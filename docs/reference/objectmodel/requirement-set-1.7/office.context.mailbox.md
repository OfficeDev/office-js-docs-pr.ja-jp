---
title: Office. メールボックス要件セット1.7
description: ''
ms.date: 04/12/2019
localization_priority: Normal
ms.openlocfilehash: 0f84e657644b198fbca722a0628a5bafcce84377
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450227"
---
# <a name="mailbox"></a><span data-ttu-id="e397a-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="e397a-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="e397a-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="e397a-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="e397a-104">Microsoft Outlook と Microsoft Outlook on the web の Outlook アドイン オブジェクト モデルへのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="e397a-104">Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="e397a-105">要件</span><span class="sxs-lookup"><span data-stu-id="e397a-105">Requirements</span></span>

|<span data-ttu-id="e397a-106">要件</span><span class="sxs-lookup"><span data-stu-id="e397a-106">Requirement</span></span>| <span data-ttu-id="e397a-107">値</span><span class="sxs-lookup"><span data-stu-id="e397a-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="e397a-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e397a-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e397a-109">1.0</span><span class="sxs-lookup"><span data-stu-id="e397a-109">1.0</span></span>|
|[<span data-ttu-id="e397a-110">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e397a-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e397a-111">制限あり</span><span class="sxs-lookup"><span data-stu-id="e397a-111">Restricted</span></span>|
|[<span data-ttu-id="e397a-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e397a-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e397a-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e397a-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="e397a-114">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="e397a-114">Members and methods</span></span>

| <span data-ttu-id="e397a-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="e397a-115">Member</span></span> | <span data-ttu-id="e397a-116">種類</span><span class="sxs-lookup"><span data-stu-id="e397a-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="e397a-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="e397a-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="e397a-118">メンバー</span><span class="sxs-lookup"><span data-stu-id="e397a-118">Member</span></span> |
| [<span data-ttu-id="e397a-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="e397a-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="e397a-120">メンバー</span><span class="sxs-lookup"><span data-stu-id="e397a-120">Member</span></span> |
| [<span data-ttu-id="e397a-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="e397a-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="e397a-122">メソッド</span><span class="sxs-lookup"><span data-stu-id="e397a-122">Method</span></span> |
| [<span data-ttu-id="e397a-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="e397a-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="e397a-124">メソッド</span><span class="sxs-lookup"><span data-stu-id="e397a-124">Method</span></span> |
| [<span data-ttu-id="e397a-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="e397a-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="e397a-126">メソッド</span><span class="sxs-lookup"><span data-stu-id="e397a-126">Method</span></span> |
| [<span data-ttu-id="e397a-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="e397a-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="e397a-128">メソッド</span><span class="sxs-lookup"><span data-stu-id="e397a-128">Method</span></span> |
| [<span data-ttu-id="e397a-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="e397a-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="e397a-130">メソッド</span><span class="sxs-lookup"><span data-stu-id="e397a-130">Method</span></span> |
| [<span data-ttu-id="e397a-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="e397a-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="e397a-132">メソッド</span><span class="sxs-lookup"><span data-stu-id="e397a-132">Method</span></span> |
| [<span data-ttu-id="e397a-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="e397a-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="e397a-134">メソッド</span><span class="sxs-lookup"><span data-stu-id="e397a-134">Method</span></span> |
| [<span data-ttu-id="e397a-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="e397a-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="e397a-136">メソッド</span><span class="sxs-lookup"><span data-stu-id="e397a-136">Method</span></span> |
| [<span data-ttu-id="e397a-137">displayNewMessageForm</span><span class="sxs-lookup"><span data-stu-id="e397a-137">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="e397a-138">メソッド</span><span class="sxs-lookup"><span data-stu-id="e397a-138">Method</span></span> |
| [<span data-ttu-id="e397a-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="e397a-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="e397a-140">メソッド</span><span class="sxs-lookup"><span data-stu-id="e397a-140">Method</span></span> |
| [<span data-ttu-id="e397a-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="e397a-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="e397a-142">メソッド</span><span class="sxs-lookup"><span data-stu-id="e397a-142">Method</span></span> |
| [<span data-ttu-id="e397a-143">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="e397a-143">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="e397a-144">メソッド</span><span class="sxs-lookup"><span data-stu-id="e397a-144">Method</span></span> |
| [<span data-ttu-id="e397a-145">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="e397a-145">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="e397a-146">メソッド</span><span class="sxs-lookup"><span data-stu-id="e397a-146">Method</span></span> |
| [<span data-ttu-id="e397a-147">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="e397a-147">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="e397a-148">メソッド</span><span class="sxs-lookup"><span data-stu-id="e397a-148">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="e397a-149">名前空間</span><span class="sxs-lookup"><span data-stu-id="e397a-149">Namespaces</span></span>

<span data-ttu-id="e397a-150">[diagnostics](Office.context.mailbox.diagnostics.md):Outlook アドインに診断情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="e397a-150">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="e397a-151">[item](Office.context.mailbox.item.md):Outlook アドインのメッセージや予定にアクセスするためのメソッドとプロパティを提供します。</span><span class="sxs-lookup"><span data-stu-id="e397a-151">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="e397a-152">[userProfile](Office.context.mailbox.userProfile.md):Outlook アドインのユーザーに関する情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="e397a-152">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="e397a-153">メンバー</span><span class="sxs-lookup"><span data-stu-id="e397a-153">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="e397a-154">ewsUrl :String</span><span class="sxs-lookup"><span data-stu-id="e397a-154">ewsUrl :String</span></span>

<span data-ttu-id="e397a-p101">このメール アカウントの Exchange Web サービス (EWS) エンドポイントの URL を取得します。読み取りモードのみです。</span><span class="sxs-lookup"><span data-stu-id="e397a-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="e397a-157">このメンバーは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="e397a-157">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="e397a-p102">
  `ewsUrl\` 値は、リモート サービスで、ユーザーのメールボックスに EWS 呼び出しを行うために使うことができます。たとえば、[選択したアイテムから添付ファイルを取得する](/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="e397a-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="e397a-160">アプリが閲覧モードで `ewsUrl` メンバーを呼び出すには、アプリのマニフェスト内に **ReadItem** アクセス許可が指定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="e397a-160">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="e397a-p103">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出してから、`ewsUrl` メンバーを使用する必要があります。アプリには、`saveAsync` メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="e397a-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="e397a-163">型</span><span class="sxs-lookup"><span data-stu-id="e397a-163">Type</span></span>

*   <span data-ttu-id="e397a-164">String</span><span class="sxs-lookup"><span data-stu-id="e397a-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e397a-165">要件</span><span class="sxs-lookup"><span data-stu-id="e397a-165">Requirements</span></span>

|<span data-ttu-id="e397a-166">要件</span><span class="sxs-lookup"><span data-stu-id="e397a-166">Requirement</span></span>| <span data-ttu-id="e397a-167">値</span><span class="sxs-lookup"><span data-stu-id="e397a-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="e397a-168">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e397a-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e397a-169">1.0</span><span class="sxs-lookup"><span data-stu-id="e397a-169">1.0</span></span>|
|[<span data-ttu-id="e397a-170">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e397a-170">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e397a-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e397a-171">ReadItem</span></span>|
|[<span data-ttu-id="e397a-172">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e397a-172">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e397a-173">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e397a-173">Compose or Read</span></span>|

---
---

#### <a name="resturl-string"></a><span data-ttu-id="e397a-174">restUrl :String</span><span class="sxs-lookup"><span data-stu-id="e397a-174">restUrl :String</span></span>

<span data-ttu-id="e397a-175">この電子メール アカウントの REST エンドポイントの URL を取得します。</span><span class="sxs-lookup"><span data-stu-id="e397a-175">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="e397a-176">`restUrl` 値は、ユーザーのメールボックスに [REST API](/outlook/rest/) 呼び出しを行うために使用できます。</span><span class="sxs-lookup"><span data-stu-id="e397a-176">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="e397a-177">アプリが閲覧モードで `restUrl` メンバーを呼び出すには、アプリのマニフェスト内に **ReadItem** アクセス許可が指定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="e397a-177">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="e397a-p104">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出してから、`restUrl` メンバーを使用する必要があります。アプリには、`saveAsync` メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="e397a-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="e397a-180">型</span><span class="sxs-lookup"><span data-stu-id="e397a-180">Type</span></span>

*   <span data-ttu-id="e397a-181">String</span><span class="sxs-lookup"><span data-stu-id="e397a-181">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e397a-182">要件</span><span class="sxs-lookup"><span data-stu-id="e397a-182">Requirements</span></span>

|<span data-ttu-id="e397a-183">要件</span><span class="sxs-lookup"><span data-stu-id="e397a-183">Requirement</span></span>| <span data-ttu-id="e397a-184">値</span><span class="sxs-lookup"><span data-stu-id="e397a-184">Value</span></span>|
|---|---|
|[<span data-ttu-id="e397a-185">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e397a-185">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e397a-186">1.5</span><span class="sxs-lookup"><span data-stu-id="e397a-186">1.5</span></span> |
|[<span data-ttu-id="e397a-187">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e397a-187">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e397a-188">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e397a-188">ReadItem</span></span>|
|[<span data-ttu-id="e397a-189">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e397a-189">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e397a-190">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e397a-190">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="e397a-191">メソッド</span><span class="sxs-lookup"><span data-stu-id="e397a-191">Methods</span></span>

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="e397a-192">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="e397a-192">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="e397a-193">サポートされているイベントのイベント ハンドラーを追加します。</span><span class="sxs-lookup"><span data-stu-id="e397a-193">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="e397a-194">現在、サポートされているイベントの種類は `Office.EventType.ItemChanged` だけです。</span><span class="sxs-lookup"><span data-stu-id="e397a-194">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e397a-195">パラメーター</span><span class="sxs-lookup"><span data-stu-id="e397a-195">Parameters</span></span>

| <span data-ttu-id="e397a-196">名前</span><span class="sxs-lookup"><span data-stu-id="e397a-196">Name</span></span> | <span data-ttu-id="e397a-197">種類</span><span class="sxs-lookup"><span data-stu-id="e397a-197">Type</span></span> | <span data-ttu-id="e397a-198">属性</span><span class="sxs-lookup"><span data-stu-id="e397a-198">Attributes</span></span> | <span data-ttu-id="e397a-199">説明</span><span class="sxs-lookup"><span data-stu-id="e397a-199">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="e397a-200">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="e397a-200">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="e397a-201">ハンドラーを呼び出す必要のあるイベント。</span><span class="sxs-lookup"><span data-stu-id="e397a-201">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="e397a-202">関数</span><span class="sxs-lookup"><span data-stu-id="e397a-202">Function</span></span> || <span data-ttu-id="e397a-p105">イベントを処理する関数。関数は、オブジェクト リテラルである単一パラメーターを受け入れる必要があります。パラメーターの `type` プロパティは、`addHandlerAsync` に渡される `eventType` パラメーターと一致します。</span><span class="sxs-lookup"><span data-stu-id="e397a-p105">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="e397a-206">Object</span><span class="sxs-lookup"><span data-stu-id="e397a-206">Object</span></span> | <span data-ttu-id="e397a-207">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="e397a-207">&lt;optional&gt;</span></span> | <span data-ttu-id="e397a-208">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="e397a-208">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="e397a-209">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="e397a-209">Object</span></span> | <span data-ttu-id="e397a-210">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="e397a-210">&lt;optional&gt;</span></span> | <span data-ttu-id="e397a-211">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="e397a-211">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="e397a-212">function</span><span class="sxs-lookup"><span data-stu-id="e397a-212">function</span></span>| <span data-ttu-id="e397a-213">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="e397a-213">&lt;optional&gt;</span></span>|<span data-ttu-id="e397a-214">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="e397a-214">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e397a-215">要件</span><span class="sxs-lookup"><span data-stu-id="e397a-215">Requirements</span></span>

|<span data-ttu-id="e397a-216">要件</span><span class="sxs-lookup"><span data-stu-id="e397a-216">Requirement</span></span>| <span data-ttu-id="e397a-217">値</span><span class="sxs-lookup"><span data-stu-id="e397a-217">Value</span></span>|
|---|---|
|[<span data-ttu-id="e397a-218">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e397a-218">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e397a-219">1.5</span><span class="sxs-lookup"><span data-stu-id="e397a-219">1.5</span></span> |
|[<span data-ttu-id="e397a-220">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e397a-220">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e397a-221">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e397a-221">ReadItem</span></span> |
|[<span data-ttu-id="e397a-222">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e397a-222">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e397a-223">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e397a-223">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e397a-224">例</span><span class="sxs-lookup"><span data-stu-id="e397a-224">Example</span></span>

```javascript
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

---
---

####  <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="e397a-225">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="e397a-225">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="e397a-226">REST 形式のアイテム ID を EWS 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="e397a-226">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="e397a-227">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="e397a-227">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="e397a-p106">REST API ([Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) や [Microsoft Graph](https://graph.microsoft.io/) など) で取得されたアイテム ID は、Exchange Web サービス (EWS) に使用される形式とは異なる形式を使用します。`convertToEwsId` メソッドは、REST 形式の ID を EWS 用の適切な形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="e397a-p106">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e397a-230">パラメーター</span><span class="sxs-lookup"><span data-stu-id="e397a-230">Parameters</span></span>

|<span data-ttu-id="e397a-231">名前</span><span class="sxs-lookup"><span data-stu-id="e397a-231">Name</span></span>| <span data-ttu-id="e397a-232">種類</span><span class="sxs-lookup"><span data-stu-id="e397a-232">Type</span></span>| <span data-ttu-id="e397a-233">説明</span><span class="sxs-lookup"><span data-stu-id="e397a-233">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="e397a-234">String</span><span class="sxs-lookup"><span data-stu-id="e397a-234">String</span></span>|<span data-ttu-id="e397a-235">Outlook REST API 形式のアイテム ID</span><span class="sxs-lookup"><span data-stu-id="e397a-235">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="e397a-236">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="e397a-236">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.restversion)|<span data-ttu-id="e397a-237">アイテム ID の取得に使用された Outlook REST API のバージョンを示す値。</span><span class="sxs-lookup"><span data-stu-id="e397a-237">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e397a-238">要件</span><span class="sxs-lookup"><span data-stu-id="e397a-238">Requirements</span></span>

|<span data-ttu-id="e397a-239">要件</span><span class="sxs-lookup"><span data-stu-id="e397a-239">Requirement</span></span>| <span data-ttu-id="e397a-240">値</span><span class="sxs-lookup"><span data-stu-id="e397a-240">Value</span></span>|
|---|---|
|[<span data-ttu-id="e397a-241">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e397a-241">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e397a-242">1.3</span><span class="sxs-lookup"><span data-stu-id="e397a-242">1.3</span></span>|
|[<span data-ttu-id="e397a-243">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e397a-243">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e397a-244">制限あり</span><span class="sxs-lookup"><span data-stu-id="e397a-244">Restricted</span></span>|
|[<span data-ttu-id="e397a-245">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e397a-245">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e397a-246">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e397a-246">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e397a-247">戻り値:</span><span class="sxs-lookup"><span data-stu-id="e397a-247">Returns:</span></span>

<span data-ttu-id="e397a-248">型:String</span><span class="sxs-lookup"><span data-stu-id="e397a-248">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="e397a-249">例</span><span class="sxs-lookup"><span data-stu-id="e397a-249">Example</span></span>

```javascript
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

---
---

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime"></a><span data-ttu-id="e397a-250">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_7/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="e397a-250">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_7/office.LocalClientTime)}</span></span>

<span data-ttu-id="e397a-251">クライアントのローカル時間で時間情報が含まれている辞書を取得します。</span><span class="sxs-lookup"><span data-stu-id="e397a-251">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="e397a-p107">Outlook 用メール アプリや Outlook Web App で使う日付と時刻では、異なるタイム ゾーンを使うことができます。Outlook では、クライアント コンピューターのタイム ゾーンを使います。Outlook Web App では、Exchange 管理センター (EAC) で設定されたタイム ゾーンを使います。ユーザー インターフェイスに表示される値が、常にユーザーが期待するタイム ゾーンと一致するように日付と時刻の値を処理する必要があります。</span><span class="sxs-lookup"><span data-stu-id="e397a-p107">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="e397a-p108">Outlook でメール アプリが実行されている場合、`convertToLocalClientTime` メソッドは、クライアント コンピューターのタイム ゾーンに設定された値のディクショナリ オブジェクトを返します。Outlook Web Apps でメール アプリが実行されている場合、`convertToLocalClientTime` メソッドは、EAC に指定されたタイム ゾーンに設定された値のディクショナリ オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="e397a-p108">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e397a-257">パラメーター</span><span class="sxs-lookup"><span data-stu-id="e397a-257">Parameters</span></span>

|<span data-ttu-id="e397a-258">名前</span><span class="sxs-lookup"><span data-stu-id="e397a-258">Name</span></span>| <span data-ttu-id="e397a-259">種類</span><span class="sxs-lookup"><span data-stu-id="e397a-259">Type</span></span>| <span data-ttu-id="e397a-260">説明</span><span class="sxs-lookup"><span data-stu-id="e397a-260">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="e397a-261">Date</span><span class="sxs-lookup"><span data-stu-id="e397a-261">Date</span></span>|<span data-ttu-id="e397a-262">日付オブジェクト</span><span class="sxs-lookup"><span data-stu-id="e397a-262">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e397a-263">要件</span><span class="sxs-lookup"><span data-stu-id="e397a-263">Requirements</span></span>

|<span data-ttu-id="e397a-264">要件</span><span class="sxs-lookup"><span data-stu-id="e397a-264">Requirement</span></span>| <span data-ttu-id="e397a-265">値</span><span class="sxs-lookup"><span data-stu-id="e397a-265">Value</span></span>|
|---|---|
|[<span data-ttu-id="e397a-266">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e397a-266">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e397a-267">1.0</span><span class="sxs-lookup"><span data-stu-id="e397a-267">1.0</span></span>|
|[<span data-ttu-id="e397a-268">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e397a-268">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e397a-269">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e397a-269">ReadItem</span></span>|
|[<span data-ttu-id="e397a-270">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e397a-270">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e397a-271">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e397a-271">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e397a-272">戻り値:</span><span class="sxs-lookup"><span data-stu-id="e397a-272">Returns:</span></span>

<span data-ttu-id="e397a-273">型:[LocalClientTime](/javascript/api/outlook_1_7/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="e397a-273">Type: [LocalClientTime](/javascript/api/outlook_1_7/office.LocalClientTime)</span></span>

---
---

####  <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="e397a-274">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="e397a-274">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="e397a-275">EWS 形式のアイテム ID を REST 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="e397a-275">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="e397a-276">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="e397a-276">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="e397a-p109">EWS または `itemId` プロパティで取得されるアイテム ID は、REST API ([Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) や [Microsoft Graph](https://graph.microsoft.io/) など) に使用される形式とは異なる形式を使用します。`convertToRestId` メソッドは、EWS 形式の ID を REST 用の適切な形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="e397a-p109">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e397a-279">パラメーター</span><span class="sxs-lookup"><span data-stu-id="e397a-279">Parameters</span></span>

|<span data-ttu-id="e397a-280">名前</span><span class="sxs-lookup"><span data-stu-id="e397a-280">Name</span></span>| <span data-ttu-id="e397a-281">種類</span><span class="sxs-lookup"><span data-stu-id="e397a-281">Type</span></span>| <span data-ttu-id="e397a-282">説明</span><span class="sxs-lookup"><span data-stu-id="e397a-282">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="e397a-283">String</span><span class="sxs-lookup"><span data-stu-id="e397a-283">String</span></span>|<span data-ttu-id="e397a-284">Exchange Web サービス (EWS) 形式のアイテム ID</span><span class="sxs-lookup"><span data-stu-id="e397a-284">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="e397a-285">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="e397a-285">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.restversion)|<span data-ttu-id="e397a-286">変換後の ID を使用する Outlook REST API のバージョンを示す値。</span><span class="sxs-lookup"><span data-stu-id="e397a-286">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e397a-287">要件</span><span class="sxs-lookup"><span data-stu-id="e397a-287">Requirements</span></span>

|<span data-ttu-id="e397a-288">要件</span><span class="sxs-lookup"><span data-stu-id="e397a-288">Requirement</span></span>| <span data-ttu-id="e397a-289">値</span><span class="sxs-lookup"><span data-stu-id="e397a-289">Value</span></span>|
|---|---|
|[<span data-ttu-id="e397a-290">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e397a-290">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e397a-291">1.3</span><span class="sxs-lookup"><span data-stu-id="e397a-291">1.3</span></span>|
|[<span data-ttu-id="e397a-292">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e397a-292">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e397a-293">制限あり</span><span class="sxs-lookup"><span data-stu-id="e397a-293">Restricted</span></span>|
|[<span data-ttu-id="e397a-294">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e397a-294">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e397a-295">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e397a-295">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e397a-296">戻り値:</span><span class="sxs-lookup"><span data-stu-id="e397a-296">Returns:</span></span>

<span data-ttu-id="e397a-297">型:String</span><span class="sxs-lookup"><span data-stu-id="e397a-297">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="e397a-298">例</span><span class="sxs-lookup"><span data-stu-id="e397a-298">Example</span></span>

```javascript
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

---
---

####  <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="e397a-299">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="e397a-299">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="e397a-300">時間情報が含まれているディクショナリから日付オブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="e397a-300">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="e397a-301">`convertToUtcClientTime` メソッドは、ローカルの日付と時刻を含むディクショナリを、ローカルの日付と時刻の正しい値を持つ日付オブジェクトに変換します。</span><span class="sxs-lookup"><span data-stu-id="e397a-301">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e397a-302">パラメーター</span><span class="sxs-lookup"><span data-stu-id="e397a-302">Parameters</span></span>

|<span data-ttu-id="e397a-303">名前</span><span class="sxs-lookup"><span data-stu-id="e397a-303">Name</span></span>| <span data-ttu-id="e397a-304">種類</span><span class="sxs-lookup"><span data-stu-id="e397a-304">Type</span></span>| <span data-ttu-id="e397a-305">説明</span><span class="sxs-lookup"><span data-stu-id="e397a-305">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="e397a-306">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="e397a-306">LocalClientTime</span></span>](/javascript/api/outlook_1_7/office.LocalClientTime)|<span data-ttu-id="e397a-307">変換するローカル時刻の値。</span><span class="sxs-lookup"><span data-stu-id="e397a-307">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e397a-308">要件</span><span class="sxs-lookup"><span data-stu-id="e397a-308">Requirements</span></span>

|<span data-ttu-id="e397a-309">要件</span><span class="sxs-lookup"><span data-stu-id="e397a-309">Requirement</span></span>| <span data-ttu-id="e397a-310">値</span><span class="sxs-lookup"><span data-stu-id="e397a-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="e397a-311">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e397a-311">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e397a-312">1.0</span><span class="sxs-lookup"><span data-stu-id="e397a-312">1.0</span></span>|
|[<span data-ttu-id="e397a-313">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e397a-313">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e397a-314">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e397a-314">ReadItem</span></span>|
|[<span data-ttu-id="e397a-315">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e397a-315">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e397a-316">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e397a-316">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e397a-317">戻り値:</span><span class="sxs-lookup"><span data-stu-id="e397a-317">Returns:</span></span>

<span data-ttu-id="e397a-318">時間が UTC で表現された日付オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="e397a-318">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="e397a-319">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="e397a-319">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="e397a-320">日付</span><span class="sxs-lookup"><span data-stu-id="e397a-320">Date</span></span></dd>

</dl>

---
---

####  <a name="displayappointmentformitemid"></a><span data-ttu-id="e397a-321">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="e397a-321">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="e397a-322">既存の予定を表示します。</span><span class="sxs-lookup"><span data-stu-id="e397a-322">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="e397a-323">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="e397a-323">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="e397a-324">`displayAppointmentForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存の予定を開きます。</span><span class="sxs-lookup"><span data-stu-id="e397a-324">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="e397a-p110">Outlook for Mac では、この方法を使って、定期的な系列の一部ではない単一の予定や定期的な系列のマスター予定を表示できます。ただし、系列のインスタンスは表示できません。これは、Outlook for Mac においては定期的な系列のインスタンスのプロパティ (アイテム ID を含む) にアクセスできないためです。</span><span class="sxs-lookup"><span data-stu-id="e397a-p110">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="e397a-327">Outlook Web App では、このメソッドは指定されたフォームの本文が 32 KB 以下の文字数の場合にフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="e397a-327">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="e397a-328">指定のアイテム識別子が既存の予定を表していない場合は、クライアント コンピューターまたはデバイスで空のウィンドウが開き、エラー メッセージは返されません。</span><span class="sxs-lookup"><span data-stu-id="e397a-328">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e397a-329">パラメーター</span><span class="sxs-lookup"><span data-stu-id="e397a-329">Parameters</span></span>

|<span data-ttu-id="e397a-330">名前</span><span class="sxs-lookup"><span data-stu-id="e397a-330">Name</span></span>| <span data-ttu-id="e397a-331">種類</span><span class="sxs-lookup"><span data-stu-id="e397a-331">Type</span></span>| <span data-ttu-id="e397a-332">説明</span><span class="sxs-lookup"><span data-stu-id="e397a-332">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="e397a-333">String</span><span class="sxs-lookup"><span data-stu-id="e397a-333">String</span></span>|<span data-ttu-id="e397a-334">既存の予定の Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="e397a-334">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e397a-335">要件</span><span class="sxs-lookup"><span data-stu-id="e397a-335">Requirements</span></span>

|<span data-ttu-id="e397a-336">要件</span><span class="sxs-lookup"><span data-stu-id="e397a-336">Requirement</span></span>| <span data-ttu-id="e397a-337">値</span><span class="sxs-lookup"><span data-stu-id="e397a-337">Value</span></span>|
|---|---|
|[<span data-ttu-id="e397a-338">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e397a-338">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e397a-339">1.0</span><span class="sxs-lookup"><span data-stu-id="e397a-339">1.0</span></span>|
|[<span data-ttu-id="e397a-340">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e397a-340">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e397a-341">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e397a-341">ReadItem</span></span>|
|[<span data-ttu-id="e397a-342">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e397a-342">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e397a-343">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e397a-343">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e397a-344">例</span><span class="sxs-lookup"><span data-stu-id="e397a-344">Example</span></span>

```javascript
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

---
---

####  <a name="displaymessageformitemid"></a><span data-ttu-id="e397a-345">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="e397a-345">displayMessageForm(itemId)</span></span>

<span data-ttu-id="e397a-346">既存のメッセージを表示します。</span><span class="sxs-lookup"><span data-stu-id="e397a-346">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="e397a-347">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="e397a-347">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="e397a-348">`displayMessageForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存のメッセージを開きます。</span><span class="sxs-lookup"><span data-stu-id="e397a-348">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="e397a-349">Outlook Web App では、このメソッドは指定されたフォームの本文が 32 KB 以下の文字数の場合にフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="e397a-349">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="e397a-350">指定のアイテム識別子が既存のメッセージを表していない場合は、クライアント コンピューターにはメッセージは表示されず、エラー メッセージも返されません。</span><span class="sxs-lookup"><span data-stu-id="e397a-350">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="e397a-p111">予定を表す `itemId` を含む `displayMessageForm` を使用しないでください。既存の予定を表示するには、`displayAppointmentForm` メソッドを使用します。新しい予定を作成するフォームを表示するには、`displayNewAppointmentForm` メソッドを使用します。</span><span class="sxs-lookup"><span data-stu-id="e397a-p111">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e397a-353">パラメーター</span><span class="sxs-lookup"><span data-stu-id="e397a-353">Parameters</span></span>

|<span data-ttu-id="e397a-354">名前</span><span class="sxs-lookup"><span data-stu-id="e397a-354">Name</span></span>| <span data-ttu-id="e397a-355">種類</span><span class="sxs-lookup"><span data-stu-id="e397a-355">Type</span></span>| <span data-ttu-id="e397a-356">説明</span><span class="sxs-lookup"><span data-stu-id="e397a-356">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="e397a-357">String</span><span class="sxs-lookup"><span data-stu-id="e397a-357">String</span></span>|<span data-ttu-id="e397a-358">既存のメッセージの Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="e397a-358">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e397a-359">要件</span><span class="sxs-lookup"><span data-stu-id="e397a-359">Requirements</span></span>

|<span data-ttu-id="e397a-360">要件</span><span class="sxs-lookup"><span data-stu-id="e397a-360">Requirement</span></span>| <span data-ttu-id="e397a-361">値</span><span class="sxs-lookup"><span data-stu-id="e397a-361">Value</span></span>|
|---|---|
|[<span data-ttu-id="e397a-362">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e397a-362">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e397a-363">1.0</span><span class="sxs-lookup"><span data-stu-id="e397a-363">1.0</span></span>|
|[<span data-ttu-id="e397a-364">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e397a-364">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e397a-365">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e397a-365">ReadItem</span></span>|
|[<span data-ttu-id="e397a-366">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e397a-366">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e397a-367">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e397a-367">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e397a-368">例</span><span class="sxs-lookup"><span data-stu-id="e397a-368">Example</span></span>

```javascript
Office.context.mailbox.displayMessageForm(messageId);
```

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="e397a-369">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="e397a-369">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="e397a-370">新しい予定を作成するためのフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="e397a-370">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="e397a-371">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="e397a-371">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="e397a-p112">`displayNewAppointmentForm` メソッドを使用すると、ユーザーが新しい予定または会議を作成できるフォームが開きます。パラメーターを指定すると、予定のフォーム フィールドにパラメーターの内容が自動的に設定されます。</span><span class="sxs-lookup"><span data-stu-id="e397a-p112">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="e397a-p113">このメソッドは、Outlook Web App と OWA for Devices において、出席者フィールドが含まれるフォームを必ず表示します。入力引数として出席者を指定しないと、このメソッドにより **[保存]** ボタンのあるフォームが表示されます。出席者を指定した場合には、フォームにその出席者と **[送信]** ボタンが表示されます。</span><span class="sxs-lookup"><span data-stu-id="e397a-p113">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="e397a-p114">Outlook リッチ クライアントと Outlook RT で、`requiredAttendees`、`optionalAttendees`、または `resources` パラメーターに出席者またはリソースを指定し、このメソッドを実行すると、**[送信]** ボタンがある会議フォームが表示されます。受信者を指定せずにこのメソッドを実行すると、**[保存して閉じる]** ボタンがある予定フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="e397a-p114">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="e397a-379">パラメーターのいずれかが指定のサイズ制限を超える場合、または不明なパラメーター名が指定されている場合は、例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="e397a-379">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e397a-380">パラメーター</span><span class="sxs-lookup"><span data-stu-id="e397a-380">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="e397a-381">すべてのパラメーターは省略可能です。</span><span class="sxs-lookup"><span data-stu-id="e397a-381">All parameters are optional.</span></span>

|<span data-ttu-id="e397a-382">名前</span><span class="sxs-lookup"><span data-stu-id="e397a-382">Name</span></span>| <span data-ttu-id="e397a-383">種類</span><span class="sxs-lookup"><span data-stu-id="e397a-383">Type</span></span>| <span data-ttu-id="e397a-384">説明</span><span class="sxs-lookup"><span data-stu-id="e397a-384">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="e397a-385">Object</span><span class="sxs-lookup"><span data-stu-id="e397a-385">Object</span></span> | <span data-ttu-id="e397a-386">新しい予定を記述するパラメーターのディクショナリ。</span><span class="sxs-lookup"><span data-stu-id="e397a-386">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="e397a-387">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="e397a-387">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="e397a-p115">予定に必要な各出席者について、メール アドレスを含む文字列の配列、または `EmailAddressDetails` オブジェクトを含む配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="e397a-p115">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="e397a-390">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="e397a-390">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="e397a-p116">予定の各任意出席者について、メール アドレスを含む文字列の配列、または `EmailAddressDetails` オブジェクトを含む配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="e397a-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="e397a-393">日付</span><span class="sxs-lookup"><span data-stu-id="e397a-393">Date</span></span> | <span data-ttu-id="e397a-394">予定の開始日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="e397a-394">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="e397a-395">日付</span><span class="sxs-lookup"><span data-stu-id="e397a-395">Date</span></span> | <span data-ttu-id="e397a-396">予定の終了日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="e397a-396">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="e397a-397">String</span><span class="sxs-lookup"><span data-stu-id="e397a-397">String</span></span> | <span data-ttu-id="e397a-p117">予定の場所を含む文字列。文字列は最大 255 文字に制限されます。</span><span class="sxs-lookup"><span data-stu-id="e397a-p117">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="e397a-400">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="e397a-400">Array.&lt;String&gt;</span></span> | <span data-ttu-id="e397a-p118">予定に必要なリソースを含む文字列の配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="e397a-p118">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="e397a-403">String</span><span class="sxs-lookup"><span data-stu-id="e397a-403">String</span></span> | <span data-ttu-id="e397a-p119">予定の件名を含む文字列です。文字列は最大 255 文字に制限されます。</span><span class="sxs-lookup"><span data-stu-id="e397a-p119">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="e397a-406">String</span><span class="sxs-lookup"><span data-stu-id="e397a-406">String</span></span> | <span data-ttu-id="e397a-p120">予定の本文。本文の内容は、最大サイズが 32 KB に制限されます。</span><span class="sxs-lookup"><span data-stu-id="e397a-p120">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="e397a-409">要件</span><span class="sxs-lookup"><span data-stu-id="e397a-409">Requirements</span></span>

|<span data-ttu-id="e397a-410">要件</span><span class="sxs-lookup"><span data-stu-id="e397a-410">Requirement</span></span>| <span data-ttu-id="e397a-411">値</span><span class="sxs-lookup"><span data-stu-id="e397a-411">Value</span></span>|
|---|---|
|[<span data-ttu-id="e397a-412">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e397a-412">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e397a-413">1.0</span><span class="sxs-lookup"><span data-stu-id="e397a-413">1.0</span></span>|
|[<span data-ttu-id="e397a-414">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e397a-414">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e397a-415">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e397a-415">ReadItem</span></span>|
|[<span data-ttu-id="e397a-416">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e397a-416">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e397a-417">読み取り</span><span class="sxs-lookup"><span data-stu-id="e397a-417">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e397a-418">例</span><span class="sxs-lookup"><span data-stu-id="e397a-418">Example</span></span>

```javascript
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

---
---

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="e397a-419">displayNewMessageForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="e397a-419">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="e397a-420">新しいメッセージを作成するためのフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="e397a-420">Displays a form for creating a new message.</span></span>

<span data-ttu-id="e397a-p121">The `displayNewMessageForm` method opens a form that enables the user to create a new message. If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span><span class="sxs-lookup"><span data-stu-id="e397a-p121">The `displayNewMessageForm` method opens a form that enables the user to create a new message. If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="e397a-423">パラメーターのいずれかが指定のサイズ制限を超える場合、または不明なパラメーター名が指定されている場合は、例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="e397a-423">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e397a-424">パラメーター</span><span class="sxs-lookup"><span data-stu-id="e397a-424">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="e397a-425">すべてのパラメーターは省略可能です。</span><span class="sxs-lookup"><span data-stu-id="e397a-425">All parameters are optional.</span></span>

|<span data-ttu-id="e397a-426">名前</span><span class="sxs-lookup"><span data-stu-id="e397a-426">Name</span></span>| <span data-ttu-id="e397a-427">種類</span><span class="sxs-lookup"><span data-stu-id="e397a-427">Type</span></span>| <span data-ttu-id="e397a-428">説明</span><span class="sxs-lookup"><span data-stu-id="e397a-428">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="e397a-429">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="e397a-429">Object</span></span> | <span data-ttu-id="e397a-430">新しいメッセージを記述するパラメータの辞書。</span><span class="sxs-lookup"><span data-stu-id="e397a-430">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="e397a-431">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="e397a-431">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="e397a-p122">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line. The array is limited to a maximum of 100 entries.</span><span class="sxs-lookup"><span data-stu-id="e397a-p122">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="e397a-434">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="e397a-434">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="e397a-p123">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line. The array is limited to a maximum of 100 entries.</span><span class="sxs-lookup"><span data-stu-id="e397a-p123">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="e397a-437">配列。&lt;文字列&gt; | 配列。&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="e397a-437">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="e397a-p124">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line. The array is limited to a maximum of 100 entries.</span><span class="sxs-lookup"><span data-stu-id="e397a-p124">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="e397a-440">String</span><span class="sxs-lookup"><span data-stu-id="e397a-440">String</span></span> | <span data-ttu-id="e397a-p125">A string containing the subject of the message. The string is limited to a maximum of 255 characters.</span><span class="sxs-lookup"><span data-stu-id="e397a-p125">A string containing the subject of the message. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="e397a-443">String</span><span class="sxs-lookup"><span data-stu-id="e397a-443">String</span></span> | <span data-ttu-id="e397a-p126">The HTML body of the message. The body content is limited to a maximum size of 32 KB.</span><span class="sxs-lookup"><span data-stu-id="e397a-p126">The HTML body of the message. The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="e397a-446">配列。&lt;オブジェクト&gt;</span><span class="sxs-lookup"><span data-stu-id="e397a-446">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="e397a-447">添付ファイルまたは添付アイテムである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="e397a-447">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="e397a-448">String</span><span class="sxs-lookup"><span data-stu-id="e397a-448">String</span></span> | <span data-ttu-id="e397a-p127">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="e397a-p127">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="e397a-451">String</span><span class="sxs-lookup"><span data-stu-id="e397a-451">String</span></span> | <span data-ttu-id="e397a-452">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="e397a-452">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="e397a-453">文字列</span><span class="sxs-lookup"><span data-stu-id="e397a-453">String</span></span> | <span data-ttu-id="e397a-p128">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="e397a-p128">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="e397a-456">ブール値</span><span class="sxs-lookup"><span data-stu-id="e397a-456">Boolean</span></span> | <span data-ttu-id="e397a-p129">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="e397a-p129">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="e397a-459">String</span><span class="sxs-lookup"><span data-stu-id="e397a-459">String</span></span> | <span data-ttu-id="e397a-p130">Only used if `type` is set to `item`. The EWS item id of the existing e-mail you want to attach to the new message. This is a string up to 100 characters.</span><span class="sxs-lookup"><span data-stu-id="e397a-p130">Only used if `type` is set to `item`. The EWS item id of the existing e-mail you want to attach to the new message. This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="e397a-463">要件</span><span class="sxs-lookup"><span data-stu-id="e397a-463">Requirements</span></span>

|<span data-ttu-id="e397a-464">要件</span><span class="sxs-lookup"><span data-stu-id="e397a-464">Requirement</span></span>| <span data-ttu-id="e397a-465">値</span><span class="sxs-lookup"><span data-stu-id="e397a-465">Value</span></span>|
|---|---|
|[<span data-ttu-id="e397a-466">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e397a-466">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e397a-467">1.6</span><span class="sxs-lookup"><span data-stu-id="e397a-467">1.6</span></span> |
|[<span data-ttu-id="e397a-468">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e397a-468">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e397a-469">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e397a-469">ReadItem</span></span>|
|[<span data-ttu-id="e397a-470">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e397a-470">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e397a-471">読み取り</span><span class="sxs-lookup"><span data-stu-id="e397a-471">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e397a-472">例</span><span class="sxs-lookup"><span data-stu-id="e397a-472">Example</span></span>

```javascript
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

---
---

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="e397a-473">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="e397a-473">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="e397a-474">REST API または Exchange Web サービスを呼び出すために使用するトークンを含む文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="e397a-474">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="e397a-p131">`getCallbackTokenAsync` メソッドは、ユーザーのメールボックスをホストする Exchange Server から不透明なトークンを取得する非同期の呼び出しを行います。コールバック トークンの有効期間は 5 分です。</span><span class="sxs-lookup"><span data-stu-id="e397a-p131">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="e397a-477">可能であれば、アドインでは Exchange Web サービスの代わりに REST API を使用することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="e397a-477">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span>

<span data-ttu-id="e397a-478">**REST トークン**</span><span class="sxs-lookup"><span data-stu-id="e397a-478">**REST Tokens**</span></span>

<span data-ttu-id="e397a-p132">REST トークンが要求された場合 (`options.isRest = true`)、結果トークンは Exchange Web サービスの呼び出しを認証するためには機能しません。アドインがマニフェストで [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) アクセス許可を指定していない限り、トークンの範囲は現在のアイテムとその添付ファイルへの読み取り専用アクセスに制限されます。`ReadWriteMailbox` アクセス許可が指定されている場合は、結果トークンは、メールを送信する機能など、メール、カレンダー、連絡先への読み取り/書き込みアクセスを付与します。</span><span class="sxs-lookup"><span data-stu-id="e397a-p132">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="e397a-482">アドインでは、`restUrl` プロパティを使用して、REST API 呼び出しを行うときに使用する正しい URL を決定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="e397a-482">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="e397a-483">**EWS トークン**</span><span class="sxs-lookup"><span data-stu-id="e397a-483">**EWS Tokens**</span></span>

<span data-ttu-id="e397a-p133">EWS トークンが要求された場合 (`options.isRest = false`)、結果トークンは REST API 呼び出しを認証するためには機能しません。トークンの範囲は、現在のアイテムへのアクセスに制限されます。</span><span class="sxs-lookup"><span data-stu-id="e397a-p133">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="e397a-486">アドインでは、`ewsUrl` プロパティを使用して、EWS 呼び出しを行うときに使用する正しい URL を決定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="e397a-486">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e397a-487">パラメーター</span><span class="sxs-lookup"><span data-stu-id="e397a-487">Parameters</span></span>

|<span data-ttu-id="e397a-488">名前</span><span class="sxs-lookup"><span data-stu-id="e397a-488">Name</span></span>| <span data-ttu-id="e397a-489">種類</span><span class="sxs-lookup"><span data-stu-id="e397a-489">Type</span></span>| <span data-ttu-id="e397a-490">属性</span><span class="sxs-lookup"><span data-stu-id="e397a-490">Attributes</span></span>| <span data-ttu-id="e397a-491">説明</span><span class="sxs-lookup"><span data-stu-id="e397a-491">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="e397a-492">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="e397a-492">Object</span></span> | <span data-ttu-id="e397a-493">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="e397a-493">&lt;optional&gt;</span></span> | <span data-ttu-id="e397a-494">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="e397a-494">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="e397a-495">Boolean</span><span class="sxs-lookup"><span data-stu-id="e397a-495">Boolean</span></span> |  <span data-ttu-id="e397a-496">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="e397a-496">&lt;optional&gt;</span></span> | <span data-ttu-id="e397a-p134">提供されたトークンを Outlook REST API または Exchange Web サービスに使用するかどうかを決定します。既定値は、`false` です。</span><span class="sxs-lookup"><span data-stu-id="e397a-p134">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="e397a-499">Object</span><span class="sxs-lookup"><span data-stu-id="e397a-499">Object</span></span> |  <span data-ttu-id="e397a-500">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="e397a-500">&lt;optional&gt;</span></span> | <span data-ttu-id="e397a-501">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="e397a-501">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="e397a-502">function</span><span class="sxs-lookup"><span data-stu-id="e397a-502">function</span></span>||<span data-ttu-id="e397a-p135">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="e397a-p135">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e397a-505">要件</span><span class="sxs-lookup"><span data-stu-id="e397a-505">Requirements</span></span>

|<span data-ttu-id="e397a-506">要件</span><span class="sxs-lookup"><span data-stu-id="e397a-506">Requirement</span></span>| <span data-ttu-id="e397a-507">値</span><span class="sxs-lookup"><span data-stu-id="e397a-507">Value</span></span>|
|---|---|
|[<span data-ttu-id="e397a-508">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e397a-508">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e397a-509">1.5</span><span class="sxs-lookup"><span data-stu-id="e397a-509">1.5</span></span> |
|[<span data-ttu-id="e397a-510">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e397a-510">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e397a-511">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e397a-511">ReadItem</span></span>|
|[<span data-ttu-id="e397a-512">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e397a-512">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e397a-513">新規作成と閲覧</span><span class="sxs-lookup"><span data-stu-id="e397a-513">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="e397a-514">例</span><span class="sxs-lookup"><span data-stu-id="e397a-514">Example</span></span>

```javascript
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

---
---

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="e397a-515">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="e397a-515">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="e397a-516">Exchange Server から添付ファイルやアイテムを取得するために使うトークンを含む文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="e397a-516">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="e397a-p136">`getCallbackTokenAsync` メソッドは、ユーザーのメールボックスをホストする Exchange Server から不透明なトークンを取得する非同期の呼び出しを行います。コールバック トークンの有効期間は 5 分です。</span><span class="sxs-lookup"><span data-stu-id="e397a-p136">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="e397a-p137">トークンと予定の識別子またはアイテムの識別子をサードパーティ システムに渡すことができます。サードパーティ システムは、トークンをベアラー承認トークンとして使用し、Exchange Web サービス (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) または [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) 操作を呼び出して、添付ファイルまたはアイテムを返します。たとえば、[選択したアイテムから添付ファイルを取得する](/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="e397a-p137">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="e397a-522">アプリが閲覧モードで `getCallbackTokenAsync` メソッドを呼び出すには、アプリのマニフェスト内に **ReadItem** アクセス許可が指定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="e397a-522">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="e397a-p138">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出してアイテムの識別子を `getCallbackTokenAsync` メソッドに渡す必要があります。アプリには、`saveAsync` メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="e397a-p138">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e397a-525">パラメーター</span><span class="sxs-lookup"><span data-stu-id="e397a-525">Parameters</span></span>

|<span data-ttu-id="e397a-526">名前</span><span class="sxs-lookup"><span data-stu-id="e397a-526">Name</span></span>| <span data-ttu-id="e397a-527">種類</span><span class="sxs-lookup"><span data-stu-id="e397a-527">Type</span></span>| <span data-ttu-id="e397a-528">属性</span><span class="sxs-lookup"><span data-stu-id="e397a-528">Attributes</span></span>| <span data-ttu-id="e397a-529">説明</span><span class="sxs-lookup"><span data-stu-id="e397a-529">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="e397a-530">function</span><span class="sxs-lookup"><span data-stu-id="e397a-530">function</span></span>||<span data-ttu-id="e397a-p139">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="e397a-p139">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="e397a-533">Object</span><span class="sxs-lookup"><span data-stu-id="e397a-533">Object</span></span>| <span data-ttu-id="e397a-534">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="e397a-534">&lt;optional&gt;</span></span>|<span data-ttu-id="e397a-535">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="e397a-535">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e397a-536">要件</span><span class="sxs-lookup"><span data-stu-id="e397a-536">Requirements</span></span>

|<span data-ttu-id="e397a-537">要件</span><span class="sxs-lookup"><span data-stu-id="e397a-537">Requirement</span></span>| <span data-ttu-id="e397a-538">値</span><span class="sxs-lookup"><span data-stu-id="e397a-538">Value</span></span>|
|---|---|
|[<span data-ttu-id="e397a-539">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e397a-539">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e397a-540">1.3</span><span class="sxs-lookup"><span data-stu-id="e397a-540">1.3</span></span>|
|[<span data-ttu-id="e397a-541">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e397a-541">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e397a-542">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e397a-542">ReadItem</span></span>|
|[<span data-ttu-id="e397a-543">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e397a-543">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e397a-544">新規作成と閲覧</span><span class="sxs-lookup"><span data-stu-id="e397a-544">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="e397a-545">例</span><span class="sxs-lookup"><span data-stu-id="e397a-545">Example</span></span>

```javascript
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

---
---

####  <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="e397a-546">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="e397a-546">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="e397a-547">ユーザーと Office アドインを識別するトークンを取得します。</span><span class="sxs-lookup"><span data-stu-id="e397a-547">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="e397a-548">`getUserIdentityTokenAsync` メソッドは、[アドインとユーザーをサード パーティのシステムで識別して認証](/outlook/add-ins/authentication)することのできるトークンを返します。</span><span class="sxs-lookup"><span data-stu-id="e397a-548">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="e397a-549">パラメーター</span><span class="sxs-lookup"><span data-stu-id="e397a-549">Parameters</span></span>

|<span data-ttu-id="e397a-550">名前</span><span class="sxs-lookup"><span data-stu-id="e397a-550">Name</span></span>| <span data-ttu-id="e397a-551">種類</span><span class="sxs-lookup"><span data-stu-id="e397a-551">Type</span></span>| <span data-ttu-id="e397a-552">属性</span><span class="sxs-lookup"><span data-stu-id="e397a-552">Attributes</span></span>| <span data-ttu-id="e397a-553">説明</span><span class="sxs-lookup"><span data-stu-id="e397a-553">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="e397a-554">関数</span><span class="sxs-lookup"><span data-stu-id="e397a-554">function</span></span>||<span data-ttu-id="e397a-555">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="e397a-555">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="e397a-556">トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="e397a-556">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="e397a-557">Object</span><span class="sxs-lookup"><span data-stu-id="e397a-557">Object</span></span>| <span data-ttu-id="e397a-558">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="e397a-558">&lt;optional&gt;</span></span>|<span data-ttu-id="e397a-559">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="e397a-559">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e397a-560">要件</span><span class="sxs-lookup"><span data-stu-id="e397a-560">Requirements</span></span>

|<span data-ttu-id="e397a-561">要件</span><span class="sxs-lookup"><span data-stu-id="e397a-561">Requirement</span></span>| <span data-ttu-id="e397a-562">値</span><span class="sxs-lookup"><span data-stu-id="e397a-562">Value</span></span>|
|---|---|
|[<span data-ttu-id="e397a-563">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e397a-563">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e397a-564">1.0</span><span class="sxs-lookup"><span data-stu-id="e397a-564">1.0</span></span>|
|[<span data-ttu-id="e397a-565">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e397a-565">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e397a-566">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e397a-566">ReadItem</span></span>|
|[<span data-ttu-id="e397a-567">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e397a-567">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e397a-568">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e397a-568">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e397a-569">例</span><span class="sxs-lookup"><span data-stu-id="e397a-569">Example</span></span>

```javascript
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

---
---

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="e397a-570">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="e397a-570">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="e397a-571">ユーザーのメールボックスをホストしている Exchange サーバー上の Exchange Web サービス (EWS) のサービスに対して非同期の要求を行います。</span><span class="sxs-lookup"><span data-stu-id="e397a-571">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="e397a-572">このメソッドは、次のシナリオではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="e397a-572">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="e397a-573">Outlook for iOS または Outlook for Android を使用している場合</span><span class="sxs-lookup"><span data-stu-id="e397a-573">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="e397a-574">アドインが Gmail のメールボックスに読み込まれる場合</span><span class="sxs-lookup"><span data-stu-id="e397a-574">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="e397a-575">このような場合は、アドインでは [REST API を使用](/outlook/add-ins/use-rest-api)して、代わりにユーザーのメールボックスにアクセスする必要があります。</span><span class="sxs-lookup"><span data-stu-id="e397a-575">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="e397a-p140">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange. See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span><span class="sxs-lookup"><span data-stu-id="e397a-p140">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange. See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="e397a-578">`makeEwsRequestAsync` メソッドでは、フォルダー関連アイテムを要求できません。</span><span class="sxs-lookup"><span data-stu-id="e397a-578">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="e397a-579">XML 要求では UTF-8 エンコードを指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="e397a-579">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="e397a-p141">`makeEwsRequestAsync` メソッドを使用するには、アドインに **ReadWriteMailbox** アクセス許可が必要です。**ReadWriteMailbox** アクセス許可と、`makeEwsRequestAsync` メソッドで呼び出せる EWS 操作の使い方については、「[ユーザーのメールボックスへのメール アドイン アクセスのアクセス許可を指定する](/outlook/add-ins/understanding-outlook-add-in-permissions)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e397a-p141">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="e397a-582">サーバー管理者は、クライアント アクセス サーバーの EWS ディレクトリで `OAuthAuthentication` を true に設定して、`makeEwsRequestAsync` メソッドで EWS 要求を行うことができるようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="e397a-582">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="e397a-583">バージョンの相違点</span><span class="sxs-lookup"><span data-stu-id="e397a-583">Version differences</span></span>

<span data-ttu-id="e397a-584">バージョン 15.0.4535.1004 より前のバージョンの Outlook で実行しているメール アプリで `makeEwsRequestAsync` メソッドを使う場合は、エンコード値を `ISO-8859-1` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="e397a-584">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="e397a-p142">Outlook on the web でメール アプリを実行している場合は、エンコード値を設定する必要はありません。mailbox.diagnostics.hostName プロパティを使って、メール アプリを Outlook で実行しているのか、Outlook on the web で実行しているのかを確認できます。mailbox.diagnostics.hostVersion プロパティを使って、どのバージョンの Outlook を使って実行しているかを確認できます。</span><span class="sxs-lookup"><span data-stu-id="e397a-p142">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e397a-588">パラメーター</span><span class="sxs-lookup"><span data-stu-id="e397a-588">Parameters</span></span>

|<span data-ttu-id="e397a-589">名前</span><span class="sxs-lookup"><span data-stu-id="e397a-589">Name</span></span>| <span data-ttu-id="e397a-590">種類</span><span class="sxs-lookup"><span data-stu-id="e397a-590">Type</span></span>| <span data-ttu-id="e397a-591">属性</span><span class="sxs-lookup"><span data-stu-id="e397a-591">Attributes</span></span>| <span data-ttu-id="e397a-592">説明</span><span class="sxs-lookup"><span data-stu-id="e397a-592">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="e397a-593">String</span><span class="sxs-lookup"><span data-stu-id="e397a-593">String</span></span>||<span data-ttu-id="e397a-594">EWS 要求です。</span><span class="sxs-lookup"><span data-stu-id="e397a-594">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="e397a-595">関数</span><span class="sxs-lookup"><span data-stu-id="e397a-595">function</span></span>||<span data-ttu-id="e397a-596">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="e397a-596">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="e397a-p143">The XML result of the EWS call is provided as a string in the `asyncResult.value` property. If the result exceeds 1 MB in size, an error message is returned instead.</span><span class="sxs-lookup"><span data-stu-id="e397a-p143">The XML result of the EWS call is provided as a string in the `asyncResult.value` property. If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="e397a-599">Object</span><span class="sxs-lookup"><span data-stu-id="e397a-599">Object</span></span>| <span data-ttu-id="e397a-600">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="e397a-600">&lt;optional&gt;</span></span>|<span data-ttu-id="e397a-601">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="e397a-601">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e397a-602">要件</span><span class="sxs-lookup"><span data-stu-id="e397a-602">Requirements</span></span>

|<span data-ttu-id="e397a-603">要件</span><span class="sxs-lookup"><span data-stu-id="e397a-603">Requirement</span></span>| <span data-ttu-id="e397a-604">値</span><span class="sxs-lookup"><span data-stu-id="e397a-604">Value</span></span>|
|---|---|
|[<span data-ttu-id="e397a-605">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e397a-605">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e397a-606">1.0</span><span class="sxs-lookup"><span data-stu-id="e397a-606">1.0</span></span>|
|[<span data-ttu-id="e397a-607">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e397a-607">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e397a-608">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="e397a-608">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="e397a-609">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e397a-609">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e397a-610">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e397a-610">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e397a-611">例</span><span class="sxs-lookup"><span data-stu-id="e397a-611">Example</span></span>

<span data-ttu-id="e397a-612">次の例は、`GetItem` 操作を使ってアイテムの件名を取得するため、`makeEwsRequestAsync` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="e397a-612">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

```javascript
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

---
---

####  <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="e397a-613">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="e397a-613">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="e397a-614">サポートされているイベントの種類のイベント ハンドラーを削除します。</span><span class="sxs-lookup"><span data-stu-id="e397a-614">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="e397a-615">現在、サポートされているイベントの種類は `Office.EventType.ItemChanged` だけです。</span><span class="sxs-lookup"><span data-stu-id="e397a-615">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e397a-616">パラメーター</span><span class="sxs-lookup"><span data-stu-id="e397a-616">Parameters</span></span>

| <span data-ttu-id="e397a-617">名前</span><span class="sxs-lookup"><span data-stu-id="e397a-617">Name</span></span> | <span data-ttu-id="e397a-618">種類</span><span class="sxs-lookup"><span data-stu-id="e397a-618">Type</span></span> | <span data-ttu-id="e397a-619">属性</span><span class="sxs-lookup"><span data-stu-id="e397a-619">Attributes</span></span> | <span data-ttu-id="e397a-620">説明</span><span class="sxs-lookup"><span data-stu-id="e397a-620">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="e397a-621">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="e397a-621">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="e397a-622">ハンドラーを取り消すイベント。</span><span class="sxs-lookup"><span data-stu-id="e397a-622">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="e397a-623">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="e397a-623">Object</span></span> | <span data-ttu-id="e397a-624">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="e397a-624">&lt;optional&gt;</span></span> | <span data-ttu-id="e397a-625">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="e397a-625">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="e397a-626">Object</span><span class="sxs-lookup"><span data-stu-id="e397a-626">Object</span></span> | <span data-ttu-id="e397a-627">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="e397a-627">&lt;optional&gt;</span></span> | <span data-ttu-id="e397a-628">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="e397a-628">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="e397a-629">function</span><span class="sxs-lookup"><span data-stu-id="e397a-629">function</span></span>| <span data-ttu-id="e397a-630">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="e397a-630">&lt;optional&gt;</span></span>|<span data-ttu-id="e397a-631">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="e397a-631">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e397a-632">要件</span><span class="sxs-lookup"><span data-stu-id="e397a-632">Requirements</span></span>

|<span data-ttu-id="e397a-633">要件</span><span class="sxs-lookup"><span data-stu-id="e397a-633">Requirement</span></span>| <span data-ttu-id="e397a-634">値</span><span class="sxs-lookup"><span data-stu-id="e397a-634">Value</span></span>|
|---|---|
|[<span data-ttu-id="e397a-635">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e397a-635">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e397a-636">1.5</span><span class="sxs-lookup"><span data-stu-id="e397a-636">1.5</span></span> |
|[<span data-ttu-id="e397a-637">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e397a-637">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e397a-638">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e397a-638">ReadItem</span></span> |
|[<span data-ttu-id="e397a-639">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e397a-639">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e397a-640">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e397a-640">Compose or Read</span></span>|
