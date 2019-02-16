---
title: Office.context.mailbox - プレビュー要件セット
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: 9df8c527b583aba8624b9cacb6115bbe001cf018
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/16/2019
ms.locfileid: "30068232"
---
# <a name="mailbox"></a><span data-ttu-id="76168-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="76168-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="76168-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="76168-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="76168-104">Microsoft Outlook と Microsoft Outlook on the web の Outlook アドイン オブジェクト モデルへのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="76168-104">Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="76168-105">要件</span><span class="sxs-lookup"><span data-stu-id="76168-105">Requirements</span></span>

|<span data-ttu-id="76168-106">要件</span><span class="sxs-lookup"><span data-stu-id="76168-106">Requirement</span></span>| <span data-ttu-id="76168-107">値</span><span class="sxs-lookup"><span data-stu-id="76168-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="76168-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="76168-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="76168-109">1.0</span><span class="sxs-lookup"><span data-stu-id="76168-109">1.0</span></span>|
|[<span data-ttu-id="76168-110">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="76168-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="76168-111">制限あり</span><span class="sxs-lookup"><span data-stu-id="76168-111">Restricted</span></span>|
|[<span data-ttu-id="76168-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="76168-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="76168-113">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="76168-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="76168-114">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="76168-114">Members and methods</span></span>

| <span data-ttu-id="76168-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="76168-115">Member</span></span> | <span data-ttu-id="76168-116">種類</span><span class="sxs-lookup"><span data-stu-id="76168-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="76168-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="76168-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="76168-118">メンバー</span><span class="sxs-lookup"><span data-stu-id="76168-118">Member</span></span> |
| [<span data-ttu-id="76168-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="76168-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="76168-120">メンバー</span><span class="sxs-lookup"><span data-stu-id="76168-120">Member</span></span> |
| [<span data-ttu-id="76168-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="76168-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="76168-122">メソッド</span><span class="sxs-lookup"><span data-stu-id="76168-122">Method</span></span> |
| [<span data-ttu-id="76168-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="76168-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="76168-124">メソッド</span><span class="sxs-lookup"><span data-stu-id="76168-124">Method</span></span> |
| [<span data-ttu-id="76168-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="76168-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) | <span data-ttu-id="76168-126">メソッド</span><span class="sxs-lookup"><span data-stu-id="76168-126">Method</span></span> |
| [<span data-ttu-id="76168-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="76168-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="76168-128">メソッド</span><span class="sxs-lookup"><span data-stu-id="76168-128">Method</span></span> |
| [<span data-ttu-id="76168-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="76168-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="76168-130">メソッド</span><span class="sxs-lookup"><span data-stu-id="76168-130">Method</span></span> |
| [<span data-ttu-id="76168-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="76168-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="76168-132">メソッド</span><span class="sxs-lookup"><span data-stu-id="76168-132">Method</span></span> |
| [<span data-ttu-id="76168-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="76168-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="76168-134">メソッド</span><span class="sxs-lookup"><span data-stu-id="76168-134">Method</span></span> |
| [<span data-ttu-id="76168-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="76168-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="76168-136">メソッド</span><span class="sxs-lookup"><span data-stu-id="76168-136">Method</span></span> |
| [<span data-ttu-id="76168-137">displayNewMessageForm</span><span class="sxs-lookup"><span data-stu-id="76168-137">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="76168-138">メソッド</span><span class="sxs-lookup"><span data-stu-id="76168-138">Method</span></span> |
| [<span data-ttu-id="76168-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="76168-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="76168-140">メソッド</span><span class="sxs-lookup"><span data-stu-id="76168-140">Method</span></span> |
| [<span data-ttu-id="76168-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="76168-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="76168-142">メソッド</span><span class="sxs-lookup"><span data-stu-id="76168-142">Method</span></span> |
| [<span data-ttu-id="76168-143">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="76168-143">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="76168-144">メソッド</span><span class="sxs-lookup"><span data-stu-id="76168-144">Method</span></span> |
| [<span data-ttu-id="76168-145">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="76168-145">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="76168-146">メソッド</span><span class="sxs-lookup"><span data-stu-id="76168-146">Method</span></span> |
| [<span data-ttu-id="76168-147">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="76168-147">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="76168-148">メソッド</span><span class="sxs-lookup"><span data-stu-id="76168-148">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="76168-149">名前空間</span><span class="sxs-lookup"><span data-stu-id="76168-149">Namespaces</span></span>

<span data-ttu-id="76168-150">[diagnostics](Office.context.mailbox.diagnostics.md):Outlook アドインに診断情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="76168-150">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="76168-151">[item](Office.context.mailbox.item.md):Outlook アドインのメッセージや予定にアクセスするためのメソッドとプロパティを提供します。</span><span class="sxs-lookup"><span data-stu-id="76168-151">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="76168-152">[userProfile](Office.context.mailbox.userProfile.md):Outlook アドインのユーザーに関する情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="76168-152">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="76168-153">メンバー</span><span class="sxs-lookup"><span data-stu-id="76168-153">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="76168-154">ewsUrl :String</span><span class="sxs-lookup"><span data-stu-id="76168-154">ewsUrl :String</span></span>

<span data-ttu-id="76168-p101">このメール アカウントの Exchange Web サービス (EWS) エンドポイントの URL を取得します。読み取りモードのみです。</span><span class="sxs-lookup"><span data-stu-id="76168-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="76168-157">このメンバーは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="76168-157">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="76168-p102">
  `ewsUrl\` 値は、リモート サービスで、ユーザーのメールボックスに EWS 呼び出しを行うために使うことができます。たとえば、[選択したアイテムから添付ファイルを取得する](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="76168-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="76168-160">アプリが閲覧モードで `ewsUrl` メンバーを呼び出すには、アプリのマニフェスト内に **ReadItem** アクセス許可が指定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="76168-160">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="76168-p103">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出してから、`ewsUrl` メンバーを使用する必要があります。アプリには、`saveAsync` メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="76168-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="76168-163">型</span><span class="sxs-lookup"><span data-stu-id="76168-163">Type</span></span>

*   <span data-ttu-id="76168-164">String</span><span class="sxs-lookup"><span data-stu-id="76168-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="76168-165">Requirements</span><span class="sxs-lookup"><span data-stu-id="76168-165">Requirements</span></span>

|<span data-ttu-id="76168-166">要件</span><span class="sxs-lookup"><span data-stu-id="76168-166">Requirement</span></span>| <span data-ttu-id="76168-167">値</span><span class="sxs-lookup"><span data-stu-id="76168-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="76168-168">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="76168-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="76168-169">1.0</span><span class="sxs-lookup"><span data-stu-id="76168-169">1.0</span></span>|
|[<span data-ttu-id="76168-170">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="76168-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="76168-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="76168-171">ReadItem</span></span>|
|[<span data-ttu-id="76168-172">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="76168-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="76168-173">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="76168-173">Compose or Read</span></span>|

#### <a name="resturl-string"></a><span data-ttu-id="76168-174">restUrl :String</span><span class="sxs-lookup"><span data-stu-id="76168-174">restUrl :String</span></span>

<span data-ttu-id="76168-175">この電子メール アカウントの REST エンドポイントの URL を取得します。</span><span class="sxs-lookup"><span data-stu-id="76168-175">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="76168-176">`restUrl` 値は、ユーザーのメールボックスに [REST API](https://docs.microsoft.com/outlook/rest/) 呼び出しを行うために使用できます。</span><span class="sxs-lookup"><span data-stu-id="76168-176">The `restUrl` value can be used to make [REST API](https://docs.microsoft.com/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="76168-177">アプリが閲覧モードで `restUrl` メンバーを呼び出すには、アプリのマニフェスト内に **ReadItem** アクセス許可が指定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="76168-177">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="76168-p104">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出してから、`restUrl` メンバーを使用する必要があります。アプリには、`saveAsync` メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="76168-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="76168-180">型</span><span class="sxs-lookup"><span data-stu-id="76168-180">Type</span></span>

*   <span data-ttu-id="76168-181">String</span><span class="sxs-lookup"><span data-stu-id="76168-181">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="76168-182">Requirements</span><span class="sxs-lookup"><span data-stu-id="76168-182">Requirements</span></span>

|<span data-ttu-id="76168-183">要件</span><span class="sxs-lookup"><span data-stu-id="76168-183">Requirement</span></span>| <span data-ttu-id="76168-184">値</span><span class="sxs-lookup"><span data-stu-id="76168-184">Value</span></span>|
|---|---|
|[<span data-ttu-id="76168-185">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="76168-185">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="76168-186">1.5</span><span class="sxs-lookup"><span data-stu-id="76168-186">1.5</span></span> |
|[<span data-ttu-id="76168-187">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="76168-187">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="76168-188">ReadItem</span><span class="sxs-lookup"><span data-stu-id="76168-188">ReadItem</span></span>|
|[<span data-ttu-id="76168-189">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="76168-189">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="76168-190">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="76168-190">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="76168-191">メソッド</span><span class="sxs-lookup"><span data-stu-id="76168-191">Methods</span></span>

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="76168-192">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="76168-192">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="76168-193">サポートされているイベントにイベント ハンドラーを追加します。</span><span class="sxs-lookup"><span data-stu-id="76168-193">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="76168-194">現在、サポートされているイベントの種類は `Office.EventType.ItemChanged` と `Office.EventType.OfficeThemeChanged` です。</span><span class="sxs-lookup"><span data-stu-id="76168-194">Currently, the supported event types are `Office.EventType.ItemChanged` and `Office.EventType.OfficeThemeChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="76168-195">パラメーター</span><span class="sxs-lookup"><span data-stu-id="76168-195">Parameters</span></span>

| <span data-ttu-id="76168-196">名前</span><span class="sxs-lookup"><span data-stu-id="76168-196">Name</span></span> | <span data-ttu-id="76168-197">型</span><span class="sxs-lookup"><span data-stu-id="76168-197">Type</span></span> | <span data-ttu-id="76168-198">属性</span><span class="sxs-lookup"><span data-stu-id="76168-198">Attributes</span></span> | <span data-ttu-id="76168-199">説明</span><span class="sxs-lookup"><span data-stu-id="76168-199">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="76168-200">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="76168-200">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="76168-201">ハンドラーを呼び出す必要のあるイベント。</span><span class="sxs-lookup"><span data-stu-id="76168-201">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="76168-202">Function</span><span class="sxs-lookup"><span data-stu-id="76168-202">Function</span></span> || <span data-ttu-id="76168-p105">イベントを処理する関数。関数は、オブジェクト リテラルである単一パラメーターを受け入れる必要があります。パラメーターの `type` プロパティは、`addHandlerAsync` に渡される `eventType` パラメーターと一致します。</span><span class="sxs-lookup"><span data-stu-id="76168-p105">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="76168-206">Object</span><span class="sxs-lookup"><span data-stu-id="76168-206">Object</span></span> | <span data-ttu-id="76168-207">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="76168-207">&lt;optional&gt;</span></span> | <span data-ttu-id="76168-208">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="76168-208">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="76168-209">Object</span><span class="sxs-lookup"><span data-stu-id="76168-209">Object</span></span> | <span data-ttu-id="76168-210">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="76168-210">&lt;optional&gt;</span></span> | <span data-ttu-id="76168-211">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="76168-211">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="76168-212">function</span><span class="sxs-lookup"><span data-stu-id="76168-212">function</span></span>| <span data-ttu-id="76168-213">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="76168-213">&lt;optional&gt;</span></span>|<span data-ttu-id="76168-214">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="76168-214">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="76168-215">要件</span><span class="sxs-lookup"><span data-stu-id="76168-215">Requirements</span></span>

|<span data-ttu-id="76168-216">要件</span><span class="sxs-lookup"><span data-stu-id="76168-216">Requirement</span></span>| <span data-ttu-id="76168-217">値</span><span class="sxs-lookup"><span data-stu-id="76168-217">Value</span></span>|
|---|---|
|[<span data-ttu-id="76168-218">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="76168-218">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="76168-219">1.5</span><span class="sxs-lookup"><span data-stu-id="76168-219">1.5</span></span> |
|[<span data-ttu-id="76168-220">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="76168-220">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="76168-221">ReadItem</span><span class="sxs-lookup"><span data-stu-id="76168-221">ReadItem</span></span> |
|[<span data-ttu-id="76168-222">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="76168-222">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="76168-223">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="76168-223">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="76168-224">例</span><span class="sxs-lookup"><span data-stu-id="76168-224">Example</span></span>

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
}
```

####  <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="76168-225">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="76168-225">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="76168-226">REST 形式のアイテム ID を EWS 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="76168-226">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="76168-227">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="76168-227">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="76168-p106">REST API ([Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) や [Microsoft Graph](https://graph.microsoft.io/) など) で取得されたアイテム ID は、Exchange Web サービス (EWS) に使用される形式とは異なる形式を使用します。`convertToEwsId` メソッドは、REST 形式の ID を EWS 用の適切な形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="76168-p106">Item IDs retrieved via a REST API (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="76168-230">パラメーター</span><span class="sxs-lookup"><span data-stu-id="76168-230">Parameters</span></span>

|<span data-ttu-id="76168-231">名前</span><span class="sxs-lookup"><span data-stu-id="76168-231">Name</span></span>| <span data-ttu-id="76168-232">種類</span><span class="sxs-lookup"><span data-stu-id="76168-232">Type</span></span>| <span data-ttu-id="76168-233">説明</span><span class="sxs-lookup"><span data-stu-id="76168-233">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="76168-234">String</span><span class="sxs-lookup"><span data-stu-id="76168-234">String</span></span>|<span data-ttu-id="76168-235">Outlook REST API 形式のアイテム ID</span><span class="sxs-lookup"><span data-stu-id="76168-235">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="76168-236">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="76168-236">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion)|<span data-ttu-id="76168-237">アイテム ID の取得に使用された Outlook REST API のバージョンを示す値。</span><span class="sxs-lookup"><span data-stu-id="76168-237">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="76168-238">要件</span><span class="sxs-lookup"><span data-stu-id="76168-238">Requirements</span></span>

|<span data-ttu-id="76168-239">要件</span><span class="sxs-lookup"><span data-stu-id="76168-239">Requirement</span></span>| <span data-ttu-id="76168-240">値</span><span class="sxs-lookup"><span data-stu-id="76168-240">Value</span></span>|
|---|---|
|[<span data-ttu-id="76168-241">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="76168-241">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="76168-242">1.3</span><span class="sxs-lookup"><span data-stu-id="76168-242">1.3</span></span>|
|[<span data-ttu-id="76168-243">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="76168-243">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="76168-244">制限あり</span><span class="sxs-lookup"><span data-stu-id="76168-244">Restricted</span></span>|
|[<span data-ttu-id="76168-245">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="76168-245">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="76168-246">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="76168-246">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="76168-247">戻り値:</span><span class="sxs-lookup"><span data-stu-id="76168-247">Returns:</span></span>

<span data-ttu-id="76168-248">型:String</span><span class="sxs-lookup"><span data-stu-id="76168-248">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="76168-249">例</span><span class="sxs-lookup"><span data-stu-id="76168-249">Example</span></span>

```javascript
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime"></a><span data-ttu-id="76168-250">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="76168-250">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime)}</span></span>

<span data-ttu-id="76168-251">クライアントのローカル時間で時間情報が含まれている辞書を取得します。</span><span class="sxs-lookup"><span data-stu-id="76168-251">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="76168-p107">Outlook 用メール アプリや Outlook Web App で使う日付と時刻では、異なるタイム ゾーンを使うことができます。Outlook では、クライアント コンピューターのタイム ゾーンを使います。Outlook Web App では、Exchange 管理センター (EAC) で設定されたタイム ゾーンを使います。ユーザー インターフェイスに表示される値が、常にユーザーが期待するタイム ゾーンと一致するように日付と時刻の値を処理する必要があります。</span><span class="sxs-lookup"><span data-stu-id="76168-p107">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="76168-p108">Outlook でメール アプリが実行されている場合、`convertToLocalClientTime` メソッドは、クライアント コンピューターのタイム ゾーンに設定された値のディクショナリ オブジェクトを返します。Outlook Web Apps でメール アプリが実行されている場合、`convertToLocalClientTime` メソッドは、EAC に指定されたタイム ゾーンに設定された値のディクショナリ オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="76168-p108">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="76168-257">パラメーター</span><span class="sxs-lookup"><span data-stu-id="76168-257">Parameters</span></span>

|<span data-ttu-id="76168-258">名前</span><span class="sxs-lookup"><span data-stu-id="76168-258">Name</span></span>| <span data-ttu-id="76168-259">型</span><span class="sxs-lookup"><span data-stu-id="76168-259">Type</span></span>| <span data-ttu-id="76168-260">説明</span><span class="sxs-lookup"><span data-stu-id="76168-260">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="76168-261">日付</span><span class="sxs-lookup"><span data-stu-id="76168-261">Date</span></span>|<span data-ttu-id="76168-262">日付オブジェクト</span><span class="sxs-lookup"><span data-stu-id="76168-262">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="76168-263">要件</span><span class="sxs-lookup"><span data-stu-id="76168-263">Requirements</span></span>

|<span data-ttu-id="76168-264">要件</span><span class="sxs-lookup"><span data-stu-id="76168-264">Requirement</span></span>| <span data-ttu-id="76168-265">値</span><span class="sxs-lookup"><span data-stu-id="76168-265">Value</span></span>|
|---|---|
|[<span data-ttu-id="76168-266">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="76168-266">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="76168-267">1.0</span><span class="sxs-lookup"><span data-stu-id="76168-267">1.0</span></span>|
|[<span data-ttu-id="76168-268">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="76168-268">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="76168-269">ReadItem</span><span class="sxs-lookup"><span data-stu-id="76168-269">ReadItem</span></span>|
|[<span data-ttu-id="76168-270">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="76168-270">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="76168-271">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="76168-271">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="76168-272">戻り値:</span><span class="sxs-lookup"><span data-stu-id="76168-272">Returns:</span></span>

<span data-ttu-id="76168-273">型:[LocalClientTime](/javascript/api/outlook/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="76168-273">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime)</span></span>

####  <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="76168-274">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="76168-274">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="76168-275">EWS 形式のアイテム ID を REST 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="76168-275">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="76168-276">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="76168-276">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="76168-p109">EWS または `itemId` プロパティで取得されるアイテム ID は、REST API ([Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) や [Microsoft Graph](https://graph.microsoft.io/) など) に使用される形式とは異なる形式を使用します。`convertToRestId` メソッドは、EWS 形式の ID を REST 用の適切な形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="76168-p109">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="76168-279">パラメーター</span><span class="sxs-lookup"><span data-stu-id="76168-279">Parameters</span></span>

|<span data-ttu-id="76168-280">名前</span><span class="sxs-lookup"><span data-stu-id="76168-280">Name</span></span>| <span data-ttu-id="76168-281">型</span><span class="sxs-lookup"><span data-stu-id="76168-281">Type</span></span>| <span data-ttu-id="76168-282">説明</span><span class="sxs-lookup"><span data-stu-id="76168-282">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="76168-283">String</span><span class="sxs-lookup"><span data-stu-id="76168-283">String</span></span>|<span data-ttu-id="76168-284">Exchange Web サービス (EWS) 形式のアイテム ID</span><span class="sxs-lookup"><span data-stu-id="76168-284">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="76168-285">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="76168-285">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion)|<span data-ttu-id="76168-286">変換後の ID を使用する Outlook REST API のバージョンを示す値。</span><span class="sxs-lookup"><span data-stu-id="76168-286">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="76168-287">要件</span><span class="sxs-lookup"><span data-stu-id="76168-287">Requirements</span></span>

|<span data-ttu-id="76168-288">要件</span><span class="sxs-lookup"><span data-stu-id="76168-288">Requirement</span></span>| <span data-ttu-id="76168-289">値</span><span class="sxs-lookup"><span data-stu-id="76168-289">Value</span></span>|
|---|---|
|[<span data-ttu-id="76168-290">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="76168-290">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="76168-291">1.3</span><span class="sxs-lookup"><span data-stu-id="76168-291">1.3</span></span>|
|[<span data-ttu-id="76168-292">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="76168-292">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="76168-293">制限あり</span><span class="sxs-lookup"><span data-stu-id="76168-293">Restricted</span></span>|
|[<span data-ttu-id="76168-294">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="76168-294">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="76168-295">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="76168-295">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="76168-296">戻り値:</span><span class="sxs-lookup"><span data-stu-id="76168-296">Returns:</span></span>

<span data-ttu-id="76168-297">型:String</span><span class="sxs-lookup"><span data-stu-id="76168-297">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="76168-298">例</span><span class="sxs-lookup"><span data-stu-id="76168-298">Example</span></span>

```javascript
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="76168-299">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="76168-299">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="76168-300">時間情報が含まれているディクショナリから日付オブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="76168-300">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="76168-301">`convertToUtcClientTime` メソッドは、ローカルの日付と時刻を含むディクショナリを、ローカルの日付と時刻の正しい値を持つ日付オブジェクトに変換します。</span><span class="sxs-lookup"><span data-stu-id="76168-301">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="76168-302">パラメーター</span><span class="sxs-lookup"><span data-stu-id="76168-302">Parameters</span></span>

|<span data-ttu-id="76168-303">名前</span><span class="sxs-lookup"><span data-stu-id="76168-303">Name</span></span>| <span data-ttu-id="76168-304">型</span><span class="sxs-lookup"><span data-stu-id="76168-304">Type</span></span>| <span data-ttu-id="76168-305">説明</span><span class="sxs-lookup"><span data-stu-id="76168-305">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="76168-306">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="76168-306">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime)|<span data-ttu-id="76168-307">変換するローカル時刻の値。</span><span class="sxs-lookup"><span data-stu-id="76168-307">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="76168-308">要件</span><span class="sxs-lookup"><span data-stu-id="76168-308">Requirements</span></span>

|<span data-ttu-id="76168-309">要件</span><span class="sxs-lookup"><span data-stu-id="76168-309">Requirement</span></span>| <span data-ttu-id="76168-310">値</span><span class="sxs-lookup"><span data-stu-id="76168-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="76168-311">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="76168-311">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="76168-312">1.0</span><span class="sxs-lookup"><span data-stu-id="76168-312">1.0</span></span>|
|[<span data-ttu-id="76168-313">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="76168-313">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="76168-314">ReadItem</span><span class="sxs-lookup"><span data-stu-id="76168-314">ReadItem</span></span>|
|[<span data-ttu-id="76168-315">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="76168-315">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="76168-316">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="76168-316">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="76168-317">戻り値:</span><span class="sxs-lookup"><span data-stu-id="76168-317">Returns:</span></span>

<span data-ttu-id="76168-318">時間が UTC で表現された日付オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="76168-318">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="76168-319">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="76168-319">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="76168-320">Date</span><span class="sxs-lookup"><span data-stu-id="76168-320">Date</span></span></dd>

</dl>

####  <a name="displayappointmentformitemid"></a><span data-ttu-id="76168-321">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="76168-321">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="76168-322">既存の予定を表示します。</span><span class="sxs-lookup"><span data-stu-id="76168-322">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="76168-323">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="76168-323">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="76168-324">`displayAppointmentForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存の予定を開きます。</span><span class="sxs-lookup"><span data-stu-id="76168-324">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="76168-p110">Outlook for Mac では、この方法を使って、定期的な系列の一部ではない単一の予定や定期的な系列のマスター予定を表示できます。ただし、系列のインスタンスは表示できません。これは、Outlook for Mac においては定期的な系列のインスタンスのプロパティ (アイテム ID を含む) にアクセスできないためです。</span><span class="sxs-lookup"><span data-stu-id="76168-p110">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="76168-327">Outlook Web App では、このメソッドは指定されたフォームの本文が 32 KB 以下の文字数の場合にフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="76168-327">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="76168-328">指定のアイテム識別子が既存の予定を表していない場合は、クライアント コンピューターまたはデバイスで空のウィンドウが開き、エラー メッセージは返されません。</span><span class="sxs-lookup"><span data-stu-id="76168-328">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="76168-329">パラメーター</span><span class="sxs-lookup"><span data-stu-id="76168-329">Parameters</span></span>

|<span data-ttu-id="76168-330">名前</span><span class="sxs-lookup"><span data-stu-id="76168-330">Name</span></span>| <span data-ttu-id="76168-331">型</span><span class="sxs-lookup"><span data-stu-id="76168-331">Type</span></span>| <span data-ttu-id="76168-332">説明</span><span class="sxs-lookup"><span data-stu-id="76168-332">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="76168-333">String</span><span class="sxs-lookup"><span data-stu-id="76168-333">String</span></span>|<span data-ttu-id="76168-334">既存の予定の Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="76168-334">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="76168-335">要件</span><span class="sxs-lookup"><span data-stu-id="76168-335">Requirements</span></span>

|<span data-ttu-id="76168-336">要件</span><span class="sxs-lookup"><span data-stu-id="76168-336">Requirement</span></span>| <span data-ttu-id="76168-337">値</span><span class="sxs-lookup"><span data-stu-id="76168-337">Value</span></span>|
|---|---|
|[<span data-ttu-id="76168-338">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="76168-338">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="76168-339">1.0</span><span class="sxs-lookup"><span data-stu-id="76168-339">1.0</span></span>|
|[<span data-ttu-id="76168-340">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="76168-340">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="76168-341">ReadItem</span><span class="sxs-lookup"><span data-stu-id="76168-341">ReadItem</span></span>|
|[<span data-ttu-id="76168-342">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="76168-342">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="76168-343">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="76168-343">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="76168-344">例</span><span class="sxs-lookup"><span data-stu-id="76168-344">Example</span></span>

```javascript
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  <a name="displaymessageformitemid"></a><span data-ttu-id="76168-345">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="76168-345">displayMessageForm(itemId)</span></span>

<span data-ttu-id="76168-346">既存のメッセージを表示します。</span><span class="sxs-lookup"><span data-stu-id="76168-346">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="76168-347">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="76168-347">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="76168-348">`displayMessageForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存のメッセージを開きます。</span><span class="sxs-lookup"><span data-stu-id="76168-348">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="76168-349">Outlook Web App では、このメソッドは指定されたフォームの本文が 32 KB 以下の文字数の場合にフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="76168-349">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="76168-350">指定のアイテム識別子が既存のメッセージを表していない場合は、クライアント コンピューターにはメッセージは表示されず、エラー メッセージも返されません。</span><span class="sxs-lookup"><span data-stu-id="76168-350">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="76168-p111">予定を表す `itemId` を含む `displayMessageForm` を使用しないでください。既存の予定を表示するには、`displayAppointmentForm` メソッドを使用します。新しい予定を作成するフォームを表示するには、`displayNewAppointmentForm` メソッドを使用します。</span><span class="sxs-lookup"><span data-stu-id="76168-p111">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="76168-353">パラメーター</span><span class="sxs-lookup"><span data-stu-id="76168-353">Parameters</span></span>

|<span data-ttu-id="76168-354">名前</span><span class="sxs-lookup"><span data-stu-id="76168-354">Name</span></span>| <span data-ttu-id="76168-355">型</span><span class="sxs-lookup"><span data-stu-id="76168-355">Type</span></span>| <span data-ttu-id="76168-356">説明</span><span class="sxs-lookup"><span data-stu-id="76168-356">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="76168-357">String</span><span class="sxs-lookup"><span data-stu-id="76168-357">String</span></span>|<span data-ttu-id="76168-358">既存のメッセージの Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="76168-358">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="76168-359">要件</span><span class="sxs-lookup"><span data-stu-id="76168-359">Requirements</span></span>

|<span data-ttu-id="76168-360">要件</span><span class="sxs-lookup"><span data-stu-id="76168-360">Requirement</span></span>| <span data-ttu-id="76168-361">値</span><span class="sxs-lookup"><span data-stu-id="76168-361">Value</span></span>|
|---|---|
|[<span data-ttu-id="76168-362">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="76168-362">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="76168-363">1.0</span><span class="sxs-lookup"><span data-stu-id="76168-363">1.0</span></span>|
|[<span data-ttu-id="76168-364">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="76168-364">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="76168-365">ReadItem</span><span class="sxs-lookup"><span data-stu-id="76168-365">ReadItem</span></span>|
|[<span data-ttu-id="76168-366">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="76168-366">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="76168-367">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="76168-367">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="76168-368">例</span><span class="sxs-lookup"><span data-stu-id="76168-368">Example</span></span>

```javascript
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="76168-369">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="76168-369">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="76168-370">新しい予定を作成するためのフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="76168-370">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="76168-371">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="76168-371">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="76168-p112">`displayNewAppointmentForm` メソッドを使用すると、ユーザーが新しい予定または会議を作成できるフォームが開きます。パラメーターを指定すると、予定のフォーム フィールドにパラメーターの内容が自動的に設定されます。</span><span class="sxs-lookup"><span data-stu-id="76168-p112">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="76168-p113">このメソッドは、Outlook Web App と OWA for Devices において、出席者フィールドが含まれるフォームを必ず表示します。入力引数として出席者を指定しないと、このメソッドにより **[保存]** ボタンのあるフォームが表示されます。出席者を指定した場合には、フォームにその出席者と **[送信]** ボタンが表示されます。</span><span class="sxs-lookup"><span data-stu-id="76168-p113">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="76168-p114">Outlook リッチ クライアントと Outlook RT で、`requiredAttendees`、`optionalAttendees`、または `resources` パラメーターに出席者またはリソースを指定し、このメソッドを実行すると、**[送信]** ボタンがある会議フォームが表示されます。受信者を指定せずにこのメソッドを実行すると、**[保存して閉じる]** ボタンがある予定フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="76168-p114">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="76168-379">パラメーターのいずれかが指定のサイズ制限を超える場合、または不明なパラメーター名が指定されている場合は、例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="76168-379">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="76168-380">パラメーター</span><span class="sxs-lookup"><span data-stu-id="76168-380">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="76168-381">すべてのパラメーターは省略可能です。</span><span class="sxs-lookup"><span data-stu-id="76168-381">All parameters are optional.</span></span>

|<span data-ttu-id="76168-382">名前</span><span class="sxs-lookup"><span data-stu-id="76168-382">Name</span></span>| <span data-ttu-id="76168-383">種類</span><span class="sxs-lookup"><span data-stu-id="76168-383">Type</span></span>| <span data-ttu-id="76168-384">説明</span><span class="sxs-lookup"><span data-stu-id="76168-384">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="76168-385">Object</span><span class="sxs-lookup"><span data-stu-id="76168-385">Object</span></span> | <span data-ttu-id="76168-386">新しい予定を記述するパラメーターのディクショナリ。</span><span class="sxs-lookup"><span data-stu-id="76168-386">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="76168-387">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="76168-387">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="76168-p115">予定に必要な各出席者について、メール アドレスを含む文字列の配列、または `EmailAddressDetails` オブジェクトを含む配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="76168-p115">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="76168-390">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="76168-390">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="76168-p116">予定の各任意出席者について、メール アドレスを含む文字列の配列、または `EmailAddressDetails` オブジェクトを含む配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="76168-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="76168-393">日付</span><span class="sxs-lookup"><span data-stu-id="76168-393">Date</span></span> | <span data-ttu-id="76168-394">予定の開始日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="76168-394">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="76168-395">Date</span><span class="sxs-lookup"><span data-stu-id="76168-395">Date</span></span> | <span data-ttu-id="76168-396">予定の終了日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="76168-396">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="76168-397">String</span><span class="sxs-lookup"><span data-stu-id="76168-397">String</span></span> | <span data-ttu-id="76168-p117">予定の場所を含む文字列。文字列は最大 255 文字に制限されます。</span><span class="sxs-lookup"><span data-stu-id="76168-p117">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="76168-400">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="76168-400">Array.&lt;String&gt;</span></span> | <span data-ttu-id="76168-p118">予定に必要なリソースを含む文字列の配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="76168-p118">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="76168-403">String</span><span class="sxs-lookup"><span data-stu-id="76168-403">String</span></span> | <span data-ttu-id="76168-p119">予定の件名を含む文字列です。文字列は最大 255 文字に制限されます。</span><span class="sxs-lookup"><span data-stu-id="76168-p119">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="76168-406">String</span><span class="sxs-lookup"><span data-stu-id="76168-406">String</span></span> | <span data-ttu-id="76168-p120">予定の本文。本文の内容は、最大サイズが 32 KB に制限されます。</span><span class="sxs-lookup"><span data-stu-id="76168-p120">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="76168-409">要件</span><span class="sxs-lookup"><span data-stu-id="76168-409">Requirements</span></span>

|<span data-ttu-id="76168-410">要件</span><span class="sxs-lookup"><span data-stu-id="76168-410">Requirement</span></span>| <span data-ttu-id="76168-411">値</span><span class="sxs-lookup"><span data-stu-id="76168-411">Value</span></span>|
|---|---|
|[<span data-ttu-id="76168-412">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="76168-412">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="76168-413">1.0</span><span class="sxs-lookup"><span data-stu-id="76168-413">1.0</span></span>|
|[<span data-ttu-id="76168-414">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="76168-414">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="76168-415">ReadItem</span><span class="sxs-lookup"><span data-stu-id="76168-415">ReadItem</span></span>|
|[<span data-ttu-id="76168-416">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="76168-416">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="76168-417">読み取り</span><span class="sxs-lookup"><span data-stu-id="76168-417">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="76168-418">例</span><span class="sxs-lookup"><span data-stu-id="76168-418">Example</span></span>

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

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="76168-419">displayNewMessageForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="76168-419">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="76168-420">新しいメッセージを作成するためのフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="76168-420">Displays a form for creating a new message.</span></span>

<span data-ttu-id="76168-421">`displayNewMessageForm` メソッドは、ユーザーが新しいメッセージを作成できるようにするフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="76168-421">The `displayNewMessageForm` method opens a form that enables the user to create a new message.</span></span> <span data-ttu-id="76168-422">パラメーターを指定すると、メッセージ フォーム フィールドにはパラメーターのコンテンツが自動的に入力されます。</span><span class="sxs-lookup"><span data-stu-id="76168-422">If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="76168-423">パラメーターのいずれかが指定のサイズ制限を超える場合、または不明なパラメーター名が指定されている場合は、例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="76168-423">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="76168-424">パラメーター</span><span class="sxs-lookup"><span data-stu-id="76168-424">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="76168-425">すべてのパラメーターは省略可能です。</span><span class="sxs-lookup"><span data-stu-id="76168-425">All parameters are optional.</span></span>

|<span data-ttu-id="76168-426">名前</span><span class="sxs-lookup"><span data-stu-id="76168-426">Name</span></span>| <span data-ttu-id="76168-427">種類</span><span class="sxs-lookup"><span data-stu-id="76168-427">Type</span></span>| <span data-ttu-id="76168-428">説明</span><span class="sxs-lookup"><span data-stu-id="76168-428">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="76168-429">Object</span><span class="sxs-lookup"><span data-stu-id="76168-429">Object</span></span> | <span data-ttu-id="76168-430">新しいメッセージを記述するパラメーターのディクショナリ。</span><span class="sxs-lookup"><span data-stu-id="76168-430">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="76168-431">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="76168-431">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="76168-432">メール アドレスを含む文字列の配列、または To 行の各受信者の `EmailAddressDetails` オブジェクトを含む配列。</span><span class="sxs-lookup"><span data-stu-id="76168-432">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line.</span></span> <span data-ttu-id="76168-433">配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="76168-433">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="76168-434">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="76168-434">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="76168-435">メール アドレスを含む文字列の配列、または Cc 行の各受信者の `EmailAddressDetails` オブジェクトを含む配列。</span><span class="sxs-lookup"><span data-stu-id="76168-435">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line.</span></span> <span data-ttu-id="76168-436">配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="76168-436">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="76168-437">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="76168-437">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="76168-438">メール アドレスを含む文字列の配列、または Bcc 行の各受信者の `EmailAddressDetails` オブジェクトを含む配列。</span><span class="sxs-lookup"><span data-stu-id="76168-438">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line.</span></span> <span data-ttu-id="76168-439">配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="76168-439">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="76168-440">String</span><span class="sxs-lookup"><span data-stu-id="76168-440">String</span></span> | <span data-ttu-id="76168-441">メッセージの件名を含む文字列。</span><span class="sxs-lookup"><span data-stu-id="76168-441">A string containing the subject of the message.</span></span> <span data-ttu-id="76168-442">文字列は最大 255 文字に制限されます。</span><span class="sxs-lookup"><span data-stu-id="76168-442">The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="76168-443">String</span><span class="sxs-lookup"><span data-stu-id="76168-443">String</span></span> | <span data-ttu-id="76168-444">メッセージの HTML 本文。</span><span class="sxs-lookup"><span data-stu-id="76168-444">The HTML body of the message.</span></span> <span data-ttu-id="76168-445">本文の内容は、最大サイズが 32 KB に制限されます。</span><span class="sxs-lookup"><span data-stu-id="76168-445">The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="76168-446">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="76168-446">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="76168-447">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="76168-447">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="76168-448">String</span><span class="sxs-lookup"><span data-stu-id="76168-448">String</span></span> | <span data-ttu-id="76168-p127">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="76168-p127">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="76168-451">String</span><span class="sxs-lookup"><span data-stu-id="76168-451">String</span></span> | <span data-ttu-id="76168-452">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="76168-452">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="76168-453">String</span><span class="sxs-lookup"><span data-stu-id="76168-453">String</span></span> | <span data-ttu-id="76168-p128">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="76168-p128">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="76168-456">Boolean</span><span class="sxs-lookup"><span data-stu-id="76168-456">Boolean</span></span> | <span data-ttu-id="76168-p129">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="76168-p129">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="76168-459">String</span><span class="sxs-lookup"><span data-stu-id="76168-459">String</span></span> | <span data-ttu-id="76168-460">`type` が `item` に設定されている場合にのみ使用されます。</span><span class="sxs-lookup"><span data-stu-id="76168-460">Only used if `type` is set to `item`.</span></span> <span data-ttu-id="76168-461">新しいメッセージに添付する必要がある既存の電子メールの EWS のアイテム ID です。</span><span class="sxs-lookup"><span data-stu-id="76168-461">The EWS item id of the existing e-mail you want to attach to the new message.</span></span> <span data-ttu-id="76168-462">最大 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="76168-462">This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="76168-463">要件</span><span class="sxs-lookup"><span data-stu-id="76168-463">Requirements</span></span>

|<span data-ttu-id="76168-464">要件</span><span class="sxs-lookup"><span data-stu-id="76168-464">Requirement</span></span>| <span data-ttu-id="76168-465">値</span><span class="sxs-lookup"><span data-stu-id="76168-465">Value</span></span>|
|---|---|
|[<span data-ttu-id="76168-466">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="76168-466">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="76168-467">1.6</span><span class="sxs-lookup"><span data-stu-id="76168-467">1.6</span></span> |
|[<span data-ttu-id="76168-468">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="76168-468">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="76168-469">ReadItem</span><span class="sxs-lookup"><span data-stu-id="76168-469">ReadItem</span></span>|
|[<span data-ttu-id="76168-470">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="76168-470">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="76168-471">読み取り</span><span class="sxs-lookup"><span data-stu-id="76168-471">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="76168-472">例</span><span class="sxs-lookup"><span data-stu-id="76168-472">Example</span></span>

```javascript
Office.context.mailbox.displayNewMessageForm(
  {
    // Copy the To line from current item.
    toRecipients: Office.context.mailbox.item.to,
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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="76168-473">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="76168-473">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="76168-474">REST API または Exchange Web サービスを呼び出すために使用するトークンを含む文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="76168-474">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="76168-p131">`getCallbackTokenAsync` メソッドは、ユーザーのメールボックスをホストする Exchange Server から不透明なトークンを取得する非同期の呼び出しを行います。コールバック トークンの有効期間は 5 分です。</span><span class="sxs-lookup"><span data-stu-id="76168-p131">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="76168-477">可能であれば、アドインでは Exchange Web サービスの代わりに REST API を使用することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="76168-477">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span>

<span data-ttu-id="76168-478">**REST トークン**</span><span class="sxs-lookup"><span data-stu-id="76168-478">**REST Tokens**</span></span>

<span data-ttu-id="76168-p132">REST トークンが要求された場合 (`options.isRest = true`)、結果トークンは Exchange Web サービスの呼び出しを認証するためには機能しません。アドインがマニフェストで [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) アクセス許可を指定していない限り、トークンの範囲は現在のアイテムとその添付ファイルへの読み取り専用アクセスに制限されます。`ReadWriteMailbox` アクセス許可が指定されている場合は、結果トークンは、メールを送信する機能など、メール、カレンダー、連絡先への読み取り/書き込みアクセスを付与します。</span><span class="sxs-lookup"><span data-stu-id="76168-p132">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="76168-482">アドインでは、`restUrl` プロパティを使用して、REST API 呼び出しを行うときに使用する正しい URL を決定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="76168-482">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="76168-483">**EWS トークン**</span><span class="sxs-lookup"><span data-stu-id="76168-483">**EWS Tokens**</span></span>

<span data-ttu-id="76168-p133">EWS トークンが要求された場合 (`options.isRest = false`)、結果トークンは REST API 呼び出しを認証するためには機能しません。トークンの範囲は、現在のアイテムへのアクセスに制限されます。</span><span class="sxs-lookup"><span data-stu-id="76168-p133">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="76168-486">アドインでは、`ewsUrl` プロパティを使用して、EWS 呼び出しを行うときに使用する正しい URL を決定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="76168-486">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="76168-487">パラメーター</span><span class="sxs-lookup"><span data-stu-id="76168-487">Parameters</span></span>

|<span data-ttu-id="76168-488">名前</span><span class="sxs-lookup"><span data-stu-id="76168-488">Name</span></span>| <span data-ttu-id="76168-489">型</span><span class="sxs-lookup"><span data-stu-id="76168-489">Type</span></span>| <span data-ttu-id="76168-490">属性</span><span class="sxs-lookup"><span data-stu-id="76168-490">Attributes</span></span>| <span data-ttu-id="76168-491">説明</span><span class="sxs-lookup"><span data-stu-id="76168-491">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="76168-492">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="76168-492">Object</span></span> | <span data-ttu-id="76168-493">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="76168-493">&lt;optional&gt;</span></span> | <span data-ttu-id="76168-494">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="76168-494">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="76168-495">Boolean</span><span class="sxs-lookup"><span data-stu-id="76168-495">Boolean</span></span> |  <span data-ttu-id="76168-496">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="76168-496">&lt;optional&gt;</span></span> | <span data-ttu-id="76168-p134">提供されたトークンを Outlook REST API または Exchange Web サービスに使用するかどうかを決定します。既定値は、`false` です。</span><span class="sxs-lookup"><span data-stu-id="76168-p134">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="76168-499">Object</span><span class="sxs-lookup"><span data-stu-id="76168-499">Object</span></span> |  <span data-ttu-id="76168-500">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="76168-500">&lt;optional&gt;</span></span> | <span data-ttu-id="76168-501">非同期メソッドに渡される状態データ。</span><span class="sxs-lookup"><span data-stu-id="76168-501">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="76168-502">function</span><span class="sxs-lookup"><span data-stu-id="76168-502">function</span></span>||<span data-ttu-id="76168-p135">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="76168-p135">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="76168-505">要件</span><span class="sxs-lookup"><span data-stu-id="76168-505">Requirements</span></span>

|<span data-ttu-id="76168-506">要件</span><span class="sxs-lookup"><span data-stu-id="76168-506">Requirement</span></span>| <span data-ttu-id="76168-507">値</span><span class="sxs-lookup"><span data-stu-id="76168-507">Value</span></span>|
|---|---|
|[<span data-ttu-id="76168-508">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="76168-508">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="76168-509">1.5</span><span class="sxs-lookup"><span data-stu-id="76168-509">1.5</span></span> |
|[<span data-ttu-id="76168-510">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="76168-510">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="76168-511">ReadItem</span><span class="sxs-lookup"><span data-stu-id="76168-511">ReadItem</span></span>|
|[<span data-ttu-id="76168-512">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="76168-512">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="76168-513">新規作成と閲覧</span><span class="sxs-lookup"><span data-stu-id="76168-513">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="76168-514">例</span><span class="sxs-lookup"><span data-stu-id="76168-514">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="76168-515">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="76168-515">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="76168-516">Exchange Server から添付ファイルやアイテムを取得するために使うトークンを含む文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="76168-516">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="76168-p136">`getCallbackTokenAsync` メソッドは、ユーザーのメールボックスをホストする Exchange Server から不透明なトークンを取得する非同期の呼び出しを行います。コールバック トークンの有効期間は 5 分です。</span><span class="sxs-lookup"><span data-stu-id="76168-p136">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="76168-p137">トークンと予定の識別子またはアイテムの識別子をサードパーティ システムに渡すことができます。サードパーティ システムは、トークンをベアラー承認トークンとして使用し、Exchange Web サービス (EWS) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) または [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) 操作を呼び出して、添付ファイルまたはアイテムを返します。たとえば、[選択したアイテムから添付ファイルを取得する](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="76168-p137">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="76168-522">アプリが閲覧モードで `getCallbackTokenAsync` メソッドを呼び出すには、アプリのマニフェスト内に **ReadItem** アクセス許可が指定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="76168-522">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="76168-p138">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出してアイテムの識別子を `getCallbackTokenAsync` メソッドに渡す必要があります。アプリには、`saveAsync` メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="76168-p138">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="76168-525">パラメーター</span><span class="sxs-lookup"><span data-stu-id="76168-525">Parameters</span></span>

|<span data-ttu-id="76168-526">名前</span><span class="sxs-lookup"><span data-stu-id="76168-526">Name</span></span>| <span data-ttu-id="76168-527">型</span><span class="sxs-lookup"><span data-stu-id="76168-527">Type</span></span>| <span data-ttu-id="76168-528">属性</span><span class="sxs-lookup"><span data-stu-id="76168-528">Attributes</span></span>| <span data-ttu-id="76168-529">説明</span><span class="sxs-lookup"><span data-stu-id="76168-529">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="76168-530">function</span><span class="sxs-lookup"><span data-stu-id="76168-530">function</span></span>||<span data-ttu-id="76168-p139">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="76168-p139">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="76168-533">Object</span><span class="sxs-lookup"><span data-stu-id="76168-533">Object</span></span>| <span data-ttu-id="76168-534">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="76168-534">&lt;optional&gt;</span></span>|<span data-ttu-id="76168-535">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="76168-535">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="76168-536">要件</span><span class="sxs-lookup"><span data-stu-id="76168-536">Requirements</span></span>

|<span data-ttu-id="76168-537">要件</span><span class="sxs-lookup"><span data-stu-id="76168-537">Requirement</span></span>| <span data-ttu-id="76168-538">値</span><span class="sxs-lookup"><span data-stu-id="76168-538">Value</span></span>|
|---|---|
|[<span data-ttu-id="76168-539">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="76168-539">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="76168-540">1.3</span><span class="sxs-lookup"><span data-stu-id="76168-540">1.3</span></span>|
|[<span data-ttu-id="76168-541">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="76168-541">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="76168-542">ReadItem</span><span class="sxs-lookup"><span data-stu-id="76168-542">ReadItem</span></span>|
|[<span data-ttu-id="76168-543">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="76168-543">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="76168-544">新規作成と閲覧</span><span class="sxs-lookup"><span data-stu-id="76168-544">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="76168-545">例</span><span class="sxs-lookup"><span data-stu-id="76168-545">Example</span></span>

```javascript
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="76168-546">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="76168-546">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="76168-547">ユーザーと Office アドインを識別するトークンを取得します。</span><span class="sxs-lookup"><span data-stu-id="76168-547">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="76168-548">`getUserIdentityTokenAsync` メソッドは、[アドインとユーザーをサード パーティのシステムで識別して認証](https://docs.microsoft.com/outlook/add-ins/authentication)することのできるトークンを返します。</span><span class="sxs-lookup"><span data-stu-id="76168-548">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](https://docs.microsoft.com/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="76168-549">パラメーター</span><span class="sxs-lookup"><span data-stu-id="76168-549">Parameters</span></span>

|<span data-ttu-id="76168-550">名前</span><span class="sxs-lookup"><span data-stu-id="76168-550">Name</span></span>| <span data-ttu-id="76168-551">型</span><span class="sxs-lookup"><span data-stu-id="76168-551">Type</span></span>| <span data-ttu-id="76168-552">属性</span><span class="sxs-lookup"><span data-stu-id="76168-552">Attributes</span></span>| <span data-ttu-id="76168-553">説明</span><span class="sxs-lookup"><span data-stu-id="76168-553">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="76168-554">関数</span><span class="sxs-lookup"><span data-stu-id="76168-554">function</span></span>||<span data-ttu-id="76168-555">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="76168-555">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="76168-556">トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="76168-556">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="76168-557">Object</span><span class="sxs-lookup"><span data-stu-id="76168-557">Object</span></span>| <span data-ttu-id="76168-558">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="76168-558">&lt;optional&gt;</span></span>|<span data-ttu-id="76168-559">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="76168-559">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="76168-560">要件</span><span class="sxs-lookup"><span data-stu-id="76168-560">Requirements</span></span>

|<span data-ttu-id="76168-561">要件</span><span class="sxs-lookup"><span data-stu-id="76168-561">Requirement</span></span>| <span data-ttu-id="76168-562">値</span><span class="sxs-lookup"><span data-stu-id="76168-562">Value</span></span>|
|---|---|
|[<span data-ttu-id="76168-563">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="76168-563">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="76168-564">1.0</span><span class="sxs-lookup"><span data-stu-id="76168-564">1.0</span></span>|
|[<span data-ttu-id="76168-565">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="76168-565">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="76168-566">ReadItem</span><span class="sxs-lookup"><span data-stu-id="76168-566">ReadItem</span></span>|
|[<span data-ttu-id="76168-567">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="76168-567">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="76168-568">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="76168-568">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="76168-569">例</span><span class="sxs-lookup"><span data-stu-id="76168-569">Example</span></span>

```javascript
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="76168-570">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="76168-570">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="76168-571">ユーザーのメールボックスをホストしている Exchange サーバー上の Exchange Web サービス (EWS) のサービスに対して非同期の要求を行います。</span><span class="sxs-lookup"><span data-stu-id="76168-571">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="76168-572">このメソッドは、次のシナリオではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="76168-572">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="76168-573">Outlook for iOS または Outlook for Android を使用している場合</span><span class="sxs-lookup"><span data-stu-id="76168-573">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="76168-574">アドインが Gmail のメールボックスに読み込まれる場合</span><span class="sxs-lookup"><span data-stu-id="76168-574">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="76168-575">このような場合は、アドインでは [REST API を使用](https://docs.microsoft.com/outlook/add-ins/use-rest-api)して、代わりにユーザーのメールボックスにアクセスする必要があります。</span><span class="sxs-lookup"><span data-stu-id="76168-575">In these cases, add-ins should [use REST APIs](https://docs.microsoft.com/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="76168-576">`makeEwsRequestAsync` メソッドは、アドインの代わりに Exchange に EWS 要求を送信します。</span><span class="sxs-lookup"><span data-stu-id="76168-576">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="76168-577">サポートされている EWS 操作の一覧については、「[Outlook アドインからの Web サービスの呼び出し](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="76168-577">See [Call web services from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="76168-578">`makeEwsRequestAsync` メソッドでは、フォルダー関連アイテムを要求できません。</span><span class="sxs-lookup"><span data-stu-id="76168-578">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="76168-579">XML 要求では UTF-8 エンコードを指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="76168-579">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="76168-p141">`makeEwsRequestAsync` メソッドを使用するには、アドインに **ReadWriteMailbox** アクセス許可が必要です。**ReadWriteMailbox** アクセス許可と、`makeEwsRequestAsync` メソッドで呼び出せる EWS 操作の使い方については、「[ユーザーのメールボックスへのメール アドイン アクセスのアクセス許可を指定する](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="76168-p141">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="76168-582">サーバー管理者は、クライアント アクセス サーバーの EWS ディレクトリで `OAuthAuthentication` を true に設定して、`makeEwsRequestAsync` メソッドで EWS 要求を行うことができるようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="76168-582">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="76168-583">バージョンの相違点</span><span class="sxs-lookup"><span data-stu-id="76168-583">Version differences</span></span>

<span data-ttu-id="76168-584">バージョン 15.0.4535.1004 より前のバージョンの Outlook で実行しているメール アプリで `makeEwsRequestAsync` メソッドを使う場合は、エンコード値を `ISO-8859-1` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="76168-584">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="76168-p142">Outlook on the web でメール アプリを実行している場合は、エンコード値を設定する必要はありません。mailbox.diagnostics.hostName プロパティを使って、メール アプリを Outlook で実行しているのか、Outlook on the web で実行しているのかを確認できます。mailbox.diagnostics.hostVersion プロパティを使って、どのバージョンの Outlook を使って実行しているかを確認できます。</span><span class="sxs-lookup"><span data-stu-id="76168-p142">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="76168-588">パラメーター</span><span class="sxs-lookup"><span data-stu-id="76168-588">Parameters</span></span>

|<span data-ttu-id="76168-589">名前</span><span class="sxs-lookup"><span data-stu-id="76168-589">Name</span></span>| <span data-ttu-id="76168-590">型</span><span class="sxs-lookup"><span data-stu-id="76168-590">Type</span></span>| <span data-ttu-id="76168-591">属性</span><span class="sxs-lookup"><span data-stu-id="76168-591">Attributes</span></span>| <span data-ttu-id="76168-592">説明</span><span class="sxs-lookup"><span data-stu-id="76168-592">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="76168-593">String</span><span class="sxs-lookup"><span data-stu-id="76168-593">String</span></span>||<span data-ttu-id="76168-594">EWS 要求です。</span><span class="sxs-lookup"><span data-stu-id="76168-594">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="76168-595">function</span><span class="sxs-lookup"><span data-stu-id="76168-595">function</span></span>||<span data-ttu-id="76168-596">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="76168-596">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="76168-597">EWS 呼び出しの XML 結果は、`asyncResult.value` プロパティ内の文字列として提供されています。</span><span class="sxs-lookup"><span data-stu-id="76168-597">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="76168-598">結果のサイズが 1 MB を超える場合、代わりにエラー メッセージが返されます。</span><span class="sxs-lookup"><span data-stu-id="76168-598">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="76168-599">Object</span><span class="sxs-lookup"><span data-stu-id="76168-599">Object</span></span>| <span data-ttu-id="76168-600">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="76168-600">&lt;optional&gt;</span></span>|<span data-ttu-id="76168-601">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="76168-601">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="76168-602">要件</span><span class="sxs-lookup"><span data-stu-id="76168-602">Requirements</span></span>

|<span data-ttu-id="76168-603">要件</span><span class="sxs-lookup"><span data-stu-id="76168-603">Requirement</span></span>| <span data-ttu-id="76168-604">値</span><span class="sxs-lookup"><span data-stu-id="76168-604">Value</span></span>|
|---|---|
|[<span data-ttu-id="76168-605">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="76168-605">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="76168-606">1.0</span><span class="sxs-lookup"><span data-stu-id="76168-606">1.0</span></span>|
|[<span data-ttu-id="76168-607">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="76168-607">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="76168-608">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="76168-608">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="76168-609">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="76168-609">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="76168-610">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="76168-610">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="76168-611">例</span><span class="sxs-lookup"><span data-stu-id="76168-611">Example</span></span>

<span data-ttu-id="76168-612">次の例は、`GetItem` 操作を使ってアイテムの件名を取得するため、`makeEwsRequestAsync` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="76168-612">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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

####  <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="76168-613">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="76168-613">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="76168-614">サポートされているイベントの種類のイベント ハンドラーを削除します。</span><span class="sxs-lookup"><span data-stu-id="76168-614">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="76168-615">現在、サポートされているイベントの種類は `Office.EventType.ItemChanged` と `Office.EventType.OfficeThemeChanged` です。</span><span class="sxs-lookup"><span data-stu-id="76168-615">Currently, the supported event types are `Office.EventType.ItemChanged` and `Office.EventType.OfficeThemeChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="76168-616">パラメーター</span><span class="sxs-lookup"><span data-stu-id="76168-616">Parameters</span></span>

| <span data-ttu-id="76168-617">名前</span><span class="sxs-lookup"><span data-stu-id="76168-617">Name</span></span> | <span data-ttu-id="76168-618">型</span><span class="sxs-lookup"><span data-stu-id="76168-618">Type</span></span> | <span data-ttu-id="76168-619">属性</span><span class="sxs-lookup"><span data-stu-id="76168-619">Attributes</span></span> | <span data-ttu-id="76168-620">説明</span><span class="sxs-lookup"><span data-stu-id="76168-620">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="76168-621">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="76168-621">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="76168-622">ハンドラーを取り消すイベント。</span><span class="sxs-lookup"><span data-stu-id="76168-622">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="76168-623">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="76168-623">Object</span></span> | <span data-ttu-id="76168-624">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="76168-624">&lt;optional&gt;</span></span> | <span data-ttu-id="76168-625">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="76168-625">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="76168-626">Object</span><span class="sxs-lookup"><span data-stu-id="76168-626">Object</span></span> | <span data-ttu-id="76168-627">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="76168-627">&lt;optional&gt;</span></span> | <span data-ttu-id="76168-628">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="76168-628">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="76168-629">function</span><span class="sxs-lookup"><span data-stu-id="76168-629">function</span></span>| <span data-ttu-id="76168-630">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="76168-630">&lt;optional&gt;</span></span>|<span data-ttu-id="76168-631">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="76168-631">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="76168-632">要件</span><span class="sxs-lookup"><span data-stu-id="76168-632">Requirements</span></span>

|<span data-ttu-id="76168-633">要件</span><span class="sxs-lookup"><span data-stu-id="76168-633">Requirement</span></span>| <span data-ttu-id="76168-634">値</span><span class="sxs-lookup"><span data-stu-id="76168-634">Value</span></span>|
|---|---|
|[<span data-ttu-id="76168-635">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="76168-635">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="76168-636">1.5</span><span class="sxs-lookup"><span data-stu-id="76168-636">1.5</span></span> |
|[<span data-ttu-id="76168-637">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="76168-637">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="76168-638">ReadItem</span><span class="sxs-lookup"><span data-stu-id="76168-638">ReadItem</span></span> |
|[<span data-ttu-id="76168-639">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="76168-639">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="76168-640">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="76168-640">Compose or Read</span></span>|
