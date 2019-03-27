---
title: Office のメールボックス-プレビュー要件セット
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 308342e041721493973015f0dad712fa4929878d
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/27/2019
ms.locfileid: "30872096"
---
# <a name="mailbox"></a><span data-ttu-id="ccbcc-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="ccbcc-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="ccbcc-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="ccbcc-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="ccbcc-104">Microsoft Outlook と Microsoft Outlook on the web の Outlook アドイン オブジェクト モデルへのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-104">Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="ccbcc-105">要件</span><span class="sxs-lookup"><span data-stu-id="ccbcc-105">Requirements</span></span>

|<span data-ttu-id="ccbcc-106">要件</span><span class="sxs-lookup"><span data-stu-id="ccbcc-106">Requirement</span></span>| <span data-ttu-id="ccbcc-107">値</span><span class="sxs-lookup"><span data-stu-id="ccbcc-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccbcc-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccbcc-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ccbcc-109">1.0</span><span class="sxs-lookup"><span data-stu-id="ccbcc-109">1.0</span></span>|
|[<span data-ttu-id="ccbcc-110">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccbcc-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ccbcc-111">制限あり</span><span class="sxs-lookup"><span data-stu-id="ccbcc-111">Restricted</span></span>|
|[<span data-ttu-id="ccbcc-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccbcc-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ccbcc-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ccbcc-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="ccbcc-114">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="ccbcc-114">Members and methods</span></span>

| <span data-ttu-id="ccbcc-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="ccbcc-115">Member</span></span> | <span data-ttu-id="ccbcc-116">種類</span><span class="sxs-lookup"><span data-stu-id="ccbcc-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="ccbcc-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="ccbcc-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="ccbcc-118">メンバー</span><span class="sxs-lookup"><span data-stu-id="ccbcc-118">Member</span></span> |
| [<span data-ttu-id="ccbcc-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="ccbcc-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="ccbcc-120">メンバー</span><span class="sxs-lookup"><span data-stu-id="ccbcc-120">Member</span></span> |
| [<span data-ttu-id="ccbcc-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="ccbcc-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="ccbcc-122">メソッド</span><span class="sxs-lookup"><span data-stu-id="ccbcc-122">Method</span></span> |
| [<span data-ttu-id="ccbcc-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="ccbcc-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="ccbcc-124">メソッド</span><span class="sxs-lookup"><span data-stu-id="ccbcc-124">Method</span></span> |
| [<span data-ttu-id="ccbcc-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="ccbcc-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="ccbcc-126">メソッド</span><span class="sxs-lookup"><span data-stu-id="ccbcc-126">Method</span></span> |
| [<span data-ttu-id="ccbcc-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="ccbcc-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="ccbcc-128">メソッド</span><span class="sxs-lookup"><span data-stu-id="ccbcc-128">Method</span></span> |
| [<span data-ttu-id="ccbcc-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="ccbcc-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="ccbcc-130">メソッド</span><span class="sxs-lookup"><span data-stu-id="ccbcc-130">Method</span></span> |
| [<span data-ttu-id="ccbcc-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="ccbcc-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="ccbcc-132">メソッド</span><span class="sxs-lookup"><span data-stu-id="ccbcc-132">Method</span></span> |
| [<span data-ttu-id="ccbcc-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="ccbcc-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="ccbcc-134">メソッド</span><span class="sxs-lookup"><span data-stu-id="ccbcc-134">Method</span></span> |
| [<span data-ttu-id="ccbcc-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="ccbcc-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="ccbcc-136">メソッド</span><span class="sxs-lookup"><span data-stu-id="ccbcc-136">Method</span></span> |
| [<span data-ttu-id="ccbcc-137">displayNewMessageForm</span><span class="sxs-lookup"><span data-stu-id="ccbcc-137">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="ccbcc-138">メソッド</span><span class="sxs-lookup"><span data-stu-id="ccbcc-138">Method</span></span> |
| [<span data-ttu-id="ccbcc-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="ccbcc-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="ccbcc-140">メソッド</span><span class="sxs-lookup"><span data-stu-id="ccbcc-140">Method</span></span> |
| [<span data-ttu-id="ccbcc-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="ccbcc-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="ccbcc-142">メソッド</span><span class="sxs-lookup"><span data-stu-id="ccbcc-142">Method</span></span> |
| [<span data-ttu-id="ccbcc-143">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="ccbcc-143">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="ccbcc-144">メソッド</span><span class="sxs-lookup"><span data-stu-id="ccbcc-144">Method</span></span> |
| [<span data-ttu-id="ccbcc-145">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="ccbcc-145">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="ccbcc-146">メソッド</span><span class="sxs-lookup"><span data-stu-id="ccbcc-146">Method</span></span> |
| [<span data-ttu-id="ccbcc-147">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="ccbcc-147">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="ccbcc-148">メソッド</span><span class="sxs-lookup"><span data-stu-id="ccbcc-148">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="ccbcc-149">名前空間</span><span class="sxs-lookup"><span data-stu-id="ccbcc-149">Namespaces</span></span>

<span data-ttu-id="ccbcc-150">[diagnostics](Office.context.mailbox.diagnostics.md): Outlook アドインに診断情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-150">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="ccbcc-151">[item](Office.context.mailbox.item.md):Outlook アドインのメッセージや予定にアクセスするためのメソッドとプロパティを提供します。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-151">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="ccbcc-152">[userProfile](Office.context.mailbox.userProfile.md):Outlook アドインのユーザーに関する情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-152">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="ccbcc-153">メンバー</span><span class="sxs-lookup"><span data-stu-id="ccbcc-153">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="ccbcc-154">ewsUrl: 文字列</span><span class="sxs-lookup"><span data-stu-id="ccbcc-154">ewsUrl :String</span></span>

<span data-ttu-id="ccbcc-p101">このメール アカウントの Exchange Web サービス (EWS) エンドポイントの URL を取得します。読み取りモードのみです。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="ccbcc-157">このメンバーは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-157">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="ccbcc-p102">`ewsUrl` 値は、リモート サービスで、ユーザーのメールボックスに EWS 呼び出しを行うために使用することができます。たとえば、[選択した項目から添付ファイルを取得する](/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="ccbcc-160">閲覧モードで `ewsUrl` メンバーを呼び出すには、アプリでマニフェスト内に **ReadItem** アクセス許可を指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-160">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="ccbcc-p103">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback)メソッドを呼び出してから、`ewsUrl`メンバーを使用する必要があります。アプリには、`saveAsync`メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="ccbcc-163">型</span><span class="sxs-lookup"><span data-stu-id="ccbcc-163">Type</span></span>

*   <span data-ttu-id="ccbcc-164">String</span><span class="sxs-lookup"><span data-stu-id="ccbcc-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ccbcc-165">要件</span><span class="sxs-lookup"><span data-stu-id="ccbcc-165">Requirements</span></span>

|<span data-ttu-id="ccbcc-166">要件</span><span class="sxs-lookup"><span data-stu-id="ccbcc-166">Requirement</span></span>| <span data-ttu-id="ccbcc-167">値</span><span class="sxs-lookup"><span data-stu-id="ccbcc-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccbcc-168">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccbcc-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ccbcc-169">1.0</span><span class="sxs-lookup"><span data-stu-id="ccbcc-169">1.0</span></span>|
|[<span data-ttu-id="ccbcc-170">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccbcc-170">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ccbcc-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ccbcc-171">ReadItem</span></span>|
|[<span data-ttu-id="ccbcc-172">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccbcc-172">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ccbcc-173">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ccbcc-173">Compose or Read</span></span>|

#### <a name="resturl-string"></a><span data-ttu-id="ccbcc-174">restUrl: 文字列</span><span class="sxs-lookup"><span data-stu-id="ccbcc-174">restUrl :String</span></span>

<span data-ttu-id="ccbcc-175">この電子メール アカウントの REST エンドポイントの URL を取得します。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-175">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="ccbcc-176">`restUrl` 値は、ユーザーのメールボックスに [REST API](/outlook/rest/) 呼び出しを行うために使用することができます。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-176">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="ccbcc-177">閲覧モードで `restUrl` メンバーを呼び出すには、アプリでマニフェスト内に **ReadItem** アクセス許可を指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-177">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="ccbcc-p104">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback)メソッドを呼び出してから、`restUrl`メンバーを使用する必要があります。アプリには、`saveAsync`メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="ccbcc-180">型</span><span class="sxs-lookup"><span data-stu-id="ccbcc-180">Type</span></span>

*   <span data-ttu-id="ccbcc-181">String</span><span class="sxs-lookup"><span data-stu-id="ccbcc-181">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ccbcc-182">要件</span><span class="sxs-lookup"><span data-stu-id="ccbcc-182">Requirements</span></span>

|<span data-ttu-id="ccbcc-183">要件</span><span class="sxs-lookup"><span data-stu-id="ccbcc-183">Requirement</span></span>| <span data-ttu-id="ccbcc-184">値</span><span class="sxs-lookup"><span data-stu-id="ccbcc-184">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccbcc-185">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccbcc-185">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ccbcc-186">1.5</span><span class="sxs-lookup"><span data-stu-id="ccbcc-186">1.5</span></span> |
|[<span data-ttu-id="ccbcc-187">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccbcc-187">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ccbcc-188">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ccbcc-188">ReadItem</span></span>|
|[<span data-ttu-id="ccbcc-189">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccbcc-189">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ccbcc-190">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ccbcc-190">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="ccbcc-191">メソッド</span><span class="sxs-lookup"><span data-stu-id="ccbcc-191">Methods</span></span>

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="ccbcc-192">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="ccbcc-192">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="ccbcc-193">サポートされているイベントのイベント ハンドラーを追加します。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-193">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="ccbcc-194">現在、サポートされている`Office.EventType.ItemChanged`イベント`Office.EventType.OfficeThemeChanged`の種類はおよびです。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-194">Currently, the supported event types are `Office.EventType.ItemChanged` and `Office.EventType.OfficeThemeChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ccbcc-195">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ccbcc-195">Parameters</span></span>

| <span data-ttu-id="ccbcc-196">名前</span><span class="sxs-lookup"><span data-stu-id="ccbcc-196">Name</span></span> | <span data-ttu-id="ccbcc-197">型</span><span class="sxs-lookup"><span data-stu-id="ccbcc-197">Type</span></span> | <span data-ttu-id="ccbcc-198">属性</span><span class="sxs-lookup"><span data-stu-id="ccbcc-198">Attributes</span></span> | <span data-ttu-id="ccbcc-199">説明</span><span class="sxs-lookup"><span data-stu-id="ccbcc-199">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="ccbcc-200">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="ccbcc-200">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="ccbcc-201">ハンドラを呼び出す必要のあるイベント。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-201">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="ccbcc-202">関数</span><span class="sxs-lookup"><span data-stu-id="ccbcc-202">Function</span></span> || <span data-ttu-id="ccbcc-p105">イベントを処理する関数。この関数は、オブジェクト リテラルである単一パラメータを受け入れる必要があります。パラメータの `type` プロパティは、`addHandlerAsync` に渡される `eventType` パラメータと一致します。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-p105">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="ccbcc-206">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="ccbcc-206">Object</span></span> | <span data-ttu-id="ccbcc-207">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="ccbcc-207">&lt;optional&gt;</span></span> | <span data-ttu-id="ccbcc-208">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-208">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="ccbcc-209">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="ccbcc-209">Object</span></span> | <span data-ttu-id="ccbcc-210">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ccbcc-210">&lt;optional&gt;</span></span> | <span data-ttu-id="ccbcc-211">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-211">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="ccbcc-212">function</span><span class="sxs-lookup"><span data-stu-id="ccbcc-212">function</span></span>| <span data-ttu-id="ccbcc-213">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ccbcc-213">&lt;optional&gt;</span></span>|<span data-ttu-id="ccbcc-214">メソッドが完了すると、`callback` パラメータに渡された関数が、単一パラメータ `asyncResult` で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-214">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ccbcc-215">要件</span><span class="sxs-lookup"><span data-stu-id="ccbcc-215">Requirements</span></span>

|<span data-ttu-id="ccbcc-216">要件</span><span class="sxs-lookup"><span data-stu-id="ccbcc-216">Requirement</span></span>| <span data-ttu-id="ccbcc-217">値</span><span class="sxs-lookup"><span data-stu-id="ccbcc-217">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccbcc-218">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccbcc-218">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ccbcc-219">1.5</span><span class="sxs-lookup"><span data-stu-id="ccbcc-219">1.5</span></span> |
|[<span data-ttu-id="ccbcc-220">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccbcc-220">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ccbcc-221">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ccbcc-221">ReadItem</span></span> |
|[<span data-ttu-id="ccbcc-222">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccbcc-222">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ccbcc-223">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ccbcc-223">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ccbcc-224">例</span><span class="sxs-lookup"><span data-stu-id="ccbcc-224">Example</span></span>

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

####  <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="ccbcc-225">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="ccbcc-225">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="ccbcc-226">REST 形式のアイテム ID を EWS 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-226">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="ccbcc-227">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-227">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="ccbcc-p106">REST API ([Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) や [Microsoft Graph](https://graph.microsoft.io/) など) 経由で取得された項目 ID は、Exchange Web サービス (EWS) で使用される形式とは異なる形式を使用します。`convertToEwsId` メソッドは、REST 形式の ID を EWS 用の適切な形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-p106">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ccbcc-230">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ccbcc-230">Parameters</span></span>

|<span data-ttu-id="ccbcc-231">名前</span><span class="sxs-lookup"><span data-stu-id="ccbcc-231">Name</span></span>| <span data-ttu-id="ccbcc-232">種類</span><span class="sxs-lookup"><span data-stu-id="ccbcc-232">Type</span></span>| <span data-ttu-id="ccbcc-233">説明</span><span class="sxs-lookup"><span data-stu-id="ccbcc-233">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="ccbcc-234">String</span><span class="sxs-lookup"><span data-stu-id="ccbcc-234">String</span></span>|<span data-ttu-id="ccbcc-235">Outlook REST API 用に書式設定された項目 ID</span><span class="sxs-lookup"><span data-stu-id="ccbcc-235">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="ccbcc-236">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="ccbcc-236">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion)|<span data-ttu-id="ccbcc-237">項目 ID の取得に使用された Outlook REST API のバージョンを示す値。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-237">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ccbcc-238">要件</span><span class="sxs-lookup"><span data-stu-id="ccbcc-238">Requirements</span></span>

|<span data-ttu-id="ccbcc-239">要件</span><span class="sxs-lookup"><span data-stu-id="ccbcc-239">Requirement</span></span>| <span data-ttu-id="ccbcc-240">値</span><span class="sxs-lookup"><span data-stu-id="ccbcc-240">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccbcc-241">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccbcc-241">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ccbcc-242">1.3</span><span class="sxs-lookup"><span data-stu-id="ccbcc-242">1.3</span></span>|
|[<span data-ttu-id="ccbcc-243">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccbcc-243">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ccbcc-244">制限あり</span><span class="sxs-lookup"><span data-stu-id="ccbcc-244">Restricted</span></span>|
|[<span data-ttu-id="ccbcc-245">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccbcc-245">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ccbcc-246">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ccbcc-246">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ccbcc-247">戻り値:</span><span class="sxs-lookup"><span data-stu-id="ccbcc-247">Returns:</span></span>

<span data-ttu-id="ccbcc-248">種類: 文字列</span><span class="sxs-lookup"><span data-stu-id="ccbcc-248">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="ccbcc-249">例</span><span class="sxs-lookup"><span data-stu-id="ccbcc-249">Example</span></span>

```javascript
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime"></a><span data-ttu-id="ccbcc-250">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="ccbcc-250">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime)}</span></span>

<span data-ttu-id="ccbcc-251">クライアントの現地時間の時間情報が含まれている辞書を取得します。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-251">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="ccbcc-p107">Outlook 用メール アプリや Outlook Web App で使う日付と時刻は、異なるタイム ゾーンを使うことができます。Outlook では、クライアント コンピューターのタイム ゾーンを使います。Outlook Web App では、Exchange 管理センター (EAC) で設定されたタイム ゾーンを使います。ユーザー インターフェイスに表示される値が、常にユーザーが期待するタイム ゾーンと一致するように日付と時刻の値を処理する必要があります。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-p107">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="ccbcc-p108">Outlook でメール アプリが実行されている場合、`convertToLocalClientTime` メソッドは、クライアント コンピューターのタイム ゾーンに設定された値で辞書オブジェクトを返します。Outlook Web Apps でメール アプリが実行されている場合、`convertToLocalClientTime` メソッドは、EAC で指定されたタイム ゾーンに設定された値で辞書オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-p108">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ccbcc-257">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ccbcc-257">Parameters</span></span>

|<span data-ttu-id="ccbcc-258">名前</span><span class="sxs-lookup"><span data-stu-id="ccbcc-258">Name</span></span>| <span data-ttu-id="ccbcc-259">種類</span><span class="sxs-lookup"><span data-stu-id="ccbcc-259">Type</span></span>| <span data-ttu-id="ccbcc-260">説明</span><span class="sxs-lookup"><span data-stu-id="ccbcc-260">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="ccbcc-261">Date</span><span class="sxs-lookup"><span data-stu-id="ccbcc-261">Date</span></span>|<span data-ttu-id="ccbcc-262">Date オブジェクト</span><span class="sxs-lookup"><span data-stu-id="ccbcc-262">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ccbcc-263">要件</span><span class="sxs-lookup"><span data-stu-id="ccbcc-263">Requirements</span></span>

|<span data-ttu-id="ccbcc-264">要件</span><span class="sxs-lookup"><span data-stu-id="ccbcc-264">Requirement</span></span>| <span data-ttu-id="ccbcc-265">値</span><span class="sxs-lookup"><span data-stu-id="ccbcc-265">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccbcc-266">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccbcc-266">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ccbcc-267">1.0</span><span class="sxs-lookup"><span data-stu-id="ccbcc-267">1.0</span></span>|
|[<span data-ttu-id="ccbcc-268">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccbcc-268">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ccbcc-269">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ccbcc-269">ReadItem</span></span>|
|[<span data-ttu-id="ccbcc-270">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccbcc-270">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ccbcc-271">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ccbcc-271">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ccbcc-272">戻り値:</span><span class="sxs-lookup"><span data-stu-id="ccbcc-272">Returns:</span></span>

<span data-ttu-id="ccbcc-273">種類:[LocalClientTime](/javascript/api/outlook/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="ccbcc-273">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime)</span></span>

####  <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="ccbcc-274">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="ccbcc-274">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="ccbcc-275">EWS 形式のアイテム ID を REST 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-275">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="ccbcc-276">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-276">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="ccbcc-p109">EWS 経由または `itemId` プロパティ経由で取得される項目 ID では、REST API ([Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) または [Microsoft Graph](https://graph.microsoft.io/) など) で使用される形式とは異なる形式を使用します。`convertToRestId` メソッドは、EWS 形式の ID を REST 用の適切な形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-p109">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ccbcc-279">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ccbcc-279">Parameters</span></span>

|<span data-ttu-id="ccbcc-280">名前</span><span class="sxs-lookup"><span data-stu-id="ccbcc-280">Name</span></span>| <span data-ttu-id="ccbcc-281">種類</span><span class="sxs-lookup"><span data-stu-id="ccbcc-281">Type</span></span>| <span data-ttu-id="ccbcc-282">説明</span><span class="sxs-lookup"><span data-stu-id="ccbcc-282">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="ccbcc-283">文字列</span><span class="sxs-lookup"><span data-stu-id="ccbcc-283">String</span></span>|<span data-ttu-id="ccbcc-284">Exchange Web サービス (EWS) 用に書式設定された項目 ID</span><span class="sxs-lookup"><span data-stu-id="ccbcc-284">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="ccbcc-285">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="ccbcc-285">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion)|<span data-ttu-id="ccbcc-286">変換後の ID とともに使用される Outlook REST API のバージョンを示す値。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-286">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ccbcc-287">要件</span><span class="sxs-lookup"><span data-stu-id="ccbcc-287">Requirements</span></span>

|<span data-ttu-id="ccbcc-288">要件</span><span class="sxs-lookup"><span data-stu-id="ccbcc-288">Requirement</span></span>| <span data-ttu-id="ccbcc-289">値</span><span class="sxs-lookup"><span data-stu-id="ccbcc-289">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccbcc-290">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccbcc-290">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ccbcc-291">1.3</span><span class="sxs-lookup"><span data-stu-id="ccbcc-291">1.3</span></span>|
|[<span data-ttu-id="ccbcc-292">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccbcc-292">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ccbcc-293">制限あり</span><span class="sxs-lookup"><span data-stu-id="ccbcc-293">Restricted</span></span>|
|[<span data-ttu-id="ccbcc-294">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccbcc-294">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ccbcc-295">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ccbcc-295">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ccbcc-296">戻り値:</span><span class="sxs-lookup"><span data-stu-id="ccbcc-296">Returns:</span></span>

<span data-ttu-id="ccbcc-297">種類: 文字列</span><span class="sxs-lookup"><span data-stu-id="ccbcc-297">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="ccbcc-298">例</span><span class="sxs-lookup"><span data-stu-id="ccbcc-298">Example</span></span>

```javascript
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="ccbcc-299">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="ccbcc-299">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="ccbcc-300">時間情報が含まれている辞書から Date オブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-300">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="ccbcc-301">`convertToUtcClientTime` メソッドは、ローカルの日付と時刻が含まれる辞書を、ローカルの日付と時刻の正しい値をもつ Date オブジェクトに変換します。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-301">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ccbcc-302">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ccbcc-302">Parameters</span></span>

|<span data-ttu-id="ccbcc-303">名前</span><span class="sxs-lookup"><span data-stu-id="ccbcc-303">Name</span></span>| <span data-ttu-id="ccbcc-304">種類</span><span class="sxs-lookup"><span data-stu-id="ccbcc-304">Type</span></span>| <span data-ttu-id="ccbcc-305">説明</span><span class="sxs-lookup"><span data-stu-id="ccbcc-305">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="ccbcc-306">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="ccbcc-306">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime)|<span data-ttu-id="ccbcc-307">変換するローカル時刻の値。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-307">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ccbcc-308">要件</span><span class="sxs-lookup"><span data-stu-id="ccbcc-308">Requirements</span></span>

|<span data-ttu-id="ccbcc-309">要件</span><span class="sxs-lookup"><span data-stu-id="ccbcc-309">Requirement</span></span>| <span data-ttu-id="ccbcc-310">値</span><span class="sxs-lookup"><span data-stu-id="ccbcc-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccbcc-311">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccbcc-311">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ccbcc-312">1.0</span><span class="sxs-lookup"><span data-stu-id="ccbcc-312">1.0</span></span>|
|[<span data-ttu-id="ccbcc-313">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccbcc-313">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ccbcc-314">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ccbcc-314">ReadItem</span></span>|
|[<span data-ttu-id="ccbcc-315">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccbcc-315">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ccbcc-316">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ccbcc-316">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ccbcc-317">戻り値 :</span><span class="sxs-lookup"><span data-stu-id="ccbcc-317">Returns:</span></span>

<span data-ttu-id="ccbcc-318">時間が UTC で表現された Date オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-318">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="ccbcc-319">

<dt>種類</dt>

</span><span class="sxs-lookup"><span data-stu-id="ccbcc-319">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="ccbcc-320">日付</span><span class="sxs-lookup"><span data-stu-id="ccbcc-320">Date</span></span></dd>

</dl>

####  <a name="displayappointmentformitemid"></a><span data-ttu-id="ccbcc-321">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="ccbcc-321">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="ccbcc-322">既存の予定を表示します。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-322">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="ccbcc-323">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-323">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="ccbcc-324">`displayAppointmentForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存の予定を開きます。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-324">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="ccbcc-p110">Outlook for Mac では、このメソッドを使用して、定期的に繰り返される予定の一部ではない単発の予定、または定期的に繰り替えされる予定の元の予定を表示できます。ただし、一連の予定のインスタンスは表示できません。これは、Outlook for Mac では、定期的に繰り返されるインスタンスのプロパティ  (項目 ID を含む) にアクセスできないためです。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-p110">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="ccbcc-327">Outlook Web App では、このメソッドで、指定されたフォームの本文が 32 KB 以下の文字数の場合に、そのフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-327">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="ccbcc-328">指定した項目識別子が既存の予定を識別しない場合には、クライアント コンピュータまたはデバイスで空白のウィンドウが開き、エラー メッセージは返されません。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-328">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ccbcc-329">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ccbcc-329">Parameters</span></span>

|<span data-ttu-id="ccbcc-330">名前</span><span class="sxs-lookup"><span data-stu-id="ccbcc-330">Name</span></span>| <span data-ttu-id="ccbcc-331">型</span><span class="sxs-lookup"><span data-stu-id="ccbcc-331">Type</span></span>| <span data-ttu-id="ccbcc-332">説明</span><span class="sxs-lookup"><span data-stu-id="ccbcc-332">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="ccbcc-333">String</span><span class="sxs-lookup"><span data-stu-id="ccbcc-333">String</span></span>|<span data-ttu-id="ccbcc-334">既存の予定表の予定の Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-334">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ccbcc-335">要件</span><span class="sxs-lookup"><span data-stu-id="ccbcc-335">Requirements</span></span>

|<span data-ttu-id="ccbcc-336">要件</span><span class="sxs-lookup"><span data-stu-id="ccbcc-336">Requirement</span></span>| <span data-ttu-id="ccbcc-337">値</span><span class="sxs-lookup"><span data-stu-id="ccbcc-337">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccbcc-338">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccbcc-338">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ccbcc-339">1.0</span><span class="sxs-lookup"><span data-stu-id="ccbcc-339">1.0</span></span>|
|[<span data-ttu-id="ccbcc-340">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccbcc-340">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ccbcc-341">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ccbcc-341">ReadItem</span></span>|
|[<span data-ttu-id="ccbcc-342">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccbcc-342">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ccbcc-343">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ccbcc-343">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ccbcc-344">例</span><span class="sxs-lookup"><span data-stu-id="ccbcc-344">Example</span></span>

```javascript
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  <a name="displaymessageformitemid"></a><span data-ttu-id="ccbcc-345">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="ccbcc-345">displayMessageForm(itemId)</span></span>

<span data-ttu-id="ccbcc-346">既存のメッセージを表示します。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-346">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="ccbcc-347">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-347">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="ccbcc-348">`displayMessageForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存のメッセージを開きます。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-348">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="ccbcc-349">Outlook Web App では、このメソッドは、指定されたフォームの本文が 32 KB 以下の文字数の場合に、そのフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-349">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="ccbcc-350">指定した項目識別子が既存のメッセージを識別しない場合には、クラアント コンピュータでメッセージは表示されず、エラー メッセージは返されません。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-350">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="ccbcc-p111">予定を表す `itemId` が含まれる `displayMessageForm` を使用しないでください。`displayAppointmentForm` メソッドを使用して既存の予定を表示し、`displayNewAppointmentForm` を使用して新しい予定を作成するフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-p111">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ccbcc-353">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ccbcc-353">Parameters</span></span>

|<span data-ttu-id="ccbcc-354">名前</span><span class="sxs-lookup"><span data-stu-id="ccbcc-354">Name</span></span>| <span data-ttu-id="ccbcc-355">型</span><span class="sxs-lookup"><span data-stu-id="ccbcc-355">Type</span></span>| <span data-ttu-id="ccbcc-356">説明</span><span class="sxs-lookup"><span data-stu-id="ccbcc-356">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="ccbcc-357">String</span><span class="sxs-lookup"><span data-stu-id="ccbcc-357">String</span></span>|<span data-ttu-id="ccbcc-358">既存のメッセージの Exchange Web サービス(EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-358">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ccbcc-359">要件</span><span class="sxs-lookup"><span data-stu-id="ccbcc-359">Requirements</span></span>

|<span data-ttu-id="ccbcc-360">要件</span><span class="sxs-lookup"><span data-stu-id="ccbcc-360">Requirement</span></span>| <span data-ttu-id="ccbcc-361">値</span><span class="sxs-lookup"><span data-stu-id="ccbcc-361">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccbcc-362">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccbcc-362">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ccbcc-363">1.0</span><span class="sxs-lookup"><span data-stu-id="ccbcc-363">1.0</span></span>|
|[<span data-ttu-id="ccbcc-364">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccbcc-364">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ccbcc-365">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ccbcc-365">ReadItem</span></span>|
|[<span data-ttu-id="ccbcc-366">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccbcc-366">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ccbcc-367">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ccbcc-367">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ccbcc-368">例</span><span class="sxs-lookup"><span data-stu-id="ccbcc-368">Example</span></span>

```javascript
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="ccbcc-369">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="ccbcc-369">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="ccbcc-370">新しい予定を作成するためのフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-370">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="ccbcc-371">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-371">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="ccbcc-p112">`displayNewAppointmentForm` メソッドを使用すると、ユーザーが新しい予定または会議を作成できるフォームが開きます。パラメータを指定すると、予定のフォーム フィールドにパラメータの内容が自動的に入力されます。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-p112">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="ccbcc-p113">Outlook Web App および OWA for Devices では、このメソッドは出席者フィールドが含まれるフォームを常に表示します。入力因数として出席者を指定しない場合には、このメソッドは [**保存**] ボタンのあるフォームを表示します。出席者を指定した場合には、フォームにはその出席者と [**送信**] ボタンが含まれます。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-p113">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="ccbcc-p114">Outlook リッチ クライアントおよび Outlook RT では、`requiredAttendees`、`optionalAttendees` または `resources` パラメータに出席者またはリソースを指定した場合、このメソッドは [**送信**] ボタンがある会議フォームを表示します。受信者を指定しない場合には、このメソッドは [**保存して閉じる**] ボタンがある予定フォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-p114">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="ccbcc-379">パラメーターのいずれかが指定のサイズ制限を超える場合、または不明なパラメーター名が指定されている場合は、例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-379">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ccbcc-380">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ccbcc-380">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="ccbcc-381">すべてのパラメーターは省略可能です。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-381">All parameters are optional.</span></span>

|<span data-ttu-id="ccbcc-382">名前</span><span class="sxs-lookup"><span data-stu-id="ccbcc-382">Name</span></span>| <span data-ttu-id="ccbcc-383">型</span><span class="sxs-lookup"><span data-stu-id="ccbcc-383">Type</span></span>| <span data-ttu-id="ccbcc-384">説明</span><span class="sxs-lookup"><span data-stu-id="ccbcc-384">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="ccbcc-385">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="ccbcc-385">Object</span></span> | <span data-ttu-id="ccbcc-386">新しい予定を記述するパラメータの辞書。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-386">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="ccbcc-387">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="ccbcc-387">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="ccbcc-p115">予定への各必須出席者の電子メール アドレスを含む文字列の配列、または出席者の `EmailAddressDetails` オブジェクトを含む配列。この配列は、最大 100 エントリまでに制限されています。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-p115">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="ccbcc-390">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="ccbcc-390">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="ccbcc-p116">予定への各任意出席者の電子メール アドレスを含む文字列の配列、または出席者の `EmailAddressDetails` オブジェクトを含む配列。この配列は、最大 100 エントリまでに制限されています。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="ccbcc-393">日付</span><span class="sxs-lookup"><span data-stu-id="ccbcc-393">Date</span></span> | <span data-ttu-id="ccbcc-394">予定の開始日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-394">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="ccbcc-395">日付</span><span class="sxs-lookup"><span data-stu-id="ccbcc-395">Date</span></span> | <span data-ttu-id="ccbcc-396">予定の終了日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-396">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="ccbcc-397">String</span><span class="sxs-lookup"><span data-stu-id="ccbcc-397">String</span></span> | <span data-ttu-id="ccbcc-p117">予定の場所を含む文字列。文字列は最大 255 文字までに制限されています。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-p117">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="ccbcc-400">配列。&lt; 文字列&gt;</span><span class="sxs-lookup"><span data-stu-id="ccbcc-400">Array.&lt;String&gt;</span></span> | <span data-ttu-id="ccbcc-p118">予定に必要なリソースを含む文字列の配列。配列は最大 100 エントリまでに制限されています。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-p118">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="ccbcc-403">String</span><span class="sxs-lookup"><span data-stu-id="ccbcc-403">String</span></span> | <span data-ttu-id="ccbcc-p119">予定の件名を含む文字列。文字列は最大 255 文字までに制限されています。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-p119">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="ccbcc-406">String</span><span class="sxs-lookup"><span data-stu-id="ccbcc-406">String</span></span> | <span data-ttu-id="ccbcc-p120">予定の本文。本文の内容は、最大サイズが 32 KB までに制限されています。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-p120">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="ccbcc-409">要件</span><span class="sxs-lookup"><span data-stu-id="ccbcc-409">Requirements</span></span>

|<span data-ttu-id="ccbcc-410">要件</span><span class="sxs-lookup"><span data-stu-id="ccbcc-410">Requirement</span></span>| <span data-ttu-id="ccbcc-411">値</span><span class="sxs-lookup"><span data-stu-id="ccbcc-411">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccbcc-412">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccbcc-412">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ccbcc-413">1.0</span><span class="sxs-lookup"><span data-stu-id="ccbcc-413">1.0</span></span>|
|[<span data-ttu-id="ccbcc-414">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccbcc-414">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ccbcc-415">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ccbcc-415">ReadItem</span></span>|
|[<span data-ttu-id="ccbcc-416">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccbcc-416">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ccbcc-417">読み取り</span><span class="sxs-lookup"><span data-stu-id="ccbcc-417">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ccbcc-418">例</span><span class="sxs-lookup"><span data-stu-id="ccbcc-418">Example</span></span>

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

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="ccbcc-419">displayNewMessageForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="ccbcc-419">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="ccbcc-420">新しいメッセージを作成するためのフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-420">Displays a form for creating a new message.</span></span>

<span data-ttu-id="ccbcc-p121">The `displayNewMessageForm` method opens a form that enables the user to create a new message. If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span><span class="sxs-lookup"><span data-stu-id="ccbcc-p121">The `displayNewMessageForm` method opens a form that enables the user to create a new message. If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="ccbcc-423">パラメータのいずれかが指定されたサイズ制限を超えた場合、または不明なパラメータ名が指定された場合には、例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-423">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ccbcc-424">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ccbcc-424">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="ccbcc-425">すべてのパラメーターは省略可能です。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-425">All parameters are optional.</span></span>

|<span data-ttu-id="ccbcc-426">名前</span><span class="sxs-lookup"><span data-stu-id="ccbcc-426">Name</span></span>| <span data-ttu-id="ccbcc-427">型</span><span class="sxs-lookup"><span data-stu-id="ccbcc-427">Type</span></span>| <span data-ttu-id="ccbcc-428">説明</span><span class="sxs-lookup"><span data-stu-id="ccbcc-428">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="ccbcc-429">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="ccbcc-429">Object</span></span> | <span data-ttu-id="ccbcc-430">新しいメッセージを記述するパラメータの辞書。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-430">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="ccbcc-431">配列。&lt;文字列&gt; | 配列。&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="ccbcc-431">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="ccbcc-p122">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line. The array is limited to a maximum of 100 entries.</span><span class="sxs-lookup"><span data-stu-id="ccbcc-p122">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="ccbcc-434">配列。&lt;文字列&gt; | 配列。&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="ccbcc-434">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="ccbcc-p123">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line. The array is limited to a maximum of 100 entries.</span><span class="sxs-lookup"><span data-stu-id="ccbcc-p123">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="ccbcc-437">配列。&lt;文字列&gt; | 配列。&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="ccbcc-437">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="ccbcc-p124">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line. The array is limited to a maximum of 100 entries.</span><span class="sxs-lookup"><span data-stu-id="ccbcc-p124">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="ccbcc-440">String</span><span class="sxs-lookup"><span data-stu-id="ccbcc-440">String</span></span> | <span data-ttu-id="ccbcc-p125">A string containing the subject of the message. The string is limited to a maximum of 255 characters.</span><span class="sxs-lookup"><span data-stu-id="ccbcc-p125">A string containing the subject of the message. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="ccbcc-443">String</span><span class="sxs-lookup"><span data-stu-id="ccbcc-443">String</span></span> | <span data-ttu-id="ccbcc-p126">The HTML body of the message. The body content is limited to a maximum size of 32 KB.</span><span class="sxs-lookup"><span data-stu-id="ccbcc-p126">The HTML body of the message. The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="ccbcc-446">配列。&lt;オブジェクト&gt;</span><span class="sxs-lookup"><span data-stu-id="ccbcc-446">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="ccbcc-447">添付ファイルまたは添付アイテムである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-447">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="ccbcc-448">String</span><span class="sxs-lookup"><span data-stu-id="ccbcc-448">String</span></span> | <span data-ttu-id="ccbcc-p127">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-p127">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="ccbcc-451">String</span><span class="sxs-lookup"><span data-stu-id="ccbcc-451">String</span></span> | <span data-ttu-id="ccbcc-452">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-452">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="ccbcc-453">文字列</span><span class="sxs-lookup"><span data-stu-id="ccbcc-453">String</span></span> | <span data-ttu-id="ccbcc-p128">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-p128">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="ccbcc-456">ブール値</span><span class="sxs-lookup"><span data-stu-id="ccbcc-456">Boolean</span></span> | <span data-ttu-id="ccbcc-p129">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-p129">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="ccbcc-459">String</span><span class="sxs-lookup"><span data-stu-id="ccbcc-459">String</span></span> | <span data-ttu-id="ccbcc-p130">Only used if `type` is set to `item`. The EWS item id of the existing e-mail you want to attach to the new message. This is a string up to 100 characters.</span><span class="sxs-lookup"><span data-stu-id="ccbcc-p130">Only used if `type` is set to `item`. The EWS item id of the existing e-mail you want to attach to the new message. This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="ccbcc-463">要件</span><span class="sxs-lookup"><span data-stu-id="ccbcc-463">Requirements</span></span>

|<span data-ttu-id="ccbcc-464">要件</span><span class="sxs-lookup"><span data-stu-id="ccbcc-464">Requirement</span></span>| <span data-ttu-id="ccbcc-465">値</span><span class="sxs-lookup"><span data-stu-id="ccbcc-465">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccbcc-466">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccbcc-466">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ccbcc-467">1.6</span><span class="sxs-lookup"><span data-stu-id="ccbcc-467">1.6</span></span> |
|[<span data-ttu-id="ccbcc-468">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccbcc-468">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ccbcc-469">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ccbcc-469">ReadItem</span></span>|
|[<span data-ttu-id="ccbcc-470">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccbcc-470">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ccbcc-471">読み取り</span><span class="sxs-lookup"><span data-stu-id="ccbcc-471">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ccbcc-472">例</span><span class="sxs-lookup"><span data-stu-id="ccbcc-472">Example</span></span>

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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="ccbcc-473">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="ccbcc-473">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="ccbcc-474">REST API または Exchange Web サービスを呼び出すために使用するトークンを含む文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-474">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="ccbcc-p131">`getCallbackTokenAsync` メソッドは、非同期の呼び出しを行なって、ユーザーのメールボックスをホストする Exchange Server から opaque トークンを取得します。コールバック トークンの有効期間は 5 分です。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-p131">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="ccbcc-477">可能であれば、アドインでは Exchange Web サービスの代わりに REST API を使用することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-477">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span>

<span data-ttu-id="ccbcc-478">**REST トークン**</span><span class="sxs-lookup"><span data-stu-id="ccbcc-478">**REST Tokens**</span></span>

<span data-ttu-id="ccbcc-p132">REST トークンが要求された場合 (`options.isRest = true`) には、作成されたトークンは Exchange Web サービスの呼び出しを認証するためには機能しません。このトークンは、アドインがマニフェストで [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) アクセス許可を指定しない限り、現在の項目およびその添付ファイルへの読み取り専用の範囲に制限されます。`ReadWriteMailbox` アクセス許可が指定された場合には、作成されるトークンは、メールを送信する機能など、メール、予定表、連絡先への読み取り/書き込みアクセスを付与します。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-p132">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="ccbcc-482">アドインでは、`restUrl`プロパティを使用して、REST API 呼び出しを行うときに使用する正しい URL を決定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-482">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="ccbcc-483">**EWS トークン**</span><span class="sxs-lookup"><span data-stu-id="ccbcc-483">**EWS Tokens**</span></span>

<span data-ttu-id="ccbcc-p133">EWS トークンが要求された場合(`options.isRest = false`) には、作成されるトークンは REST API の呼び出しを認証するためには機能しません。このトークンは、現在の項目にアクセスできる範囲に制限されます。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-p133">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="ccbcc-486">アドインでは、`ewsUrl` プロパティを使用して、EWS 呼び出しを行うときに使用する正しい URL を決定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-486">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ccbcc-487">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ccbcc-487">Parameters</span></span>

|<span data-ttu-id="ccbcc-488">名前</span><span class="sxs-lookup"><span data-stu-id="ccbcc-488">Name</span></span>| <span data-ttu-id="ccbcc-489">型</span><span class="sxs-lookup"><span data-stu-id="ccbcc-489">Type</span></span>| <span data-ttu-id="ccbcc-490">属性</span><span class="sxs-lookup"><span data-stu-id="ccbcc-490">Attributes</span></span>| <span data-ttu-id="ccbcc-491">説明</span><span class="sxs-lookup"><span data-stu-id="ccbcc-491">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="ccbcc-492">Object</span><span class="sxs-lookup"><span data-stu-id="ccbcc-492">Object</span></span> | <span data-ttu-id="ccbcc-493">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="ccbcc-493">&lt;optional&gt;</span></span> | <span data-ttu-id="ccbcc-494">次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-494">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="ccbcc-495">ブール値</span><span class="sxs-lookup"><span data-stu-id="ccbcc-495">Boolean</span></span> |  <span data-ttu-id="ccbcc-496">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ccbcc-496">&lt;optional&gt;</span></span> | <span data-ttu-id="ccbcc-p134">提供されたトークンを Outlook REST API または Exchange Web サービスに使用するかどうかを決定します。既定値は、`false`です。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-p134">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="ccbcc-499">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="ccbcc-499">Object</span></span> |  <span data-ttu-id="ccbcc-500">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="ccbcc-500">&lt;optional&gt;</span></span> | <span data-ttu-id="ccbcc-501">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-501">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="ccbcc-502">function</span><span class="sxs-lookup"><span data-stu-id="ccbcc-502">function</span></span>||<span data-ttu-id="ccbcc-p135">メソッドが完了すると、`callback` パラメータに渡された関数が、単一パラメータ `asyncResult` で呼び出されます。このパラメータは、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトです。トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-p135">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ccbcc-505">要件</span><span class="sxs-lookup"><span data-stu-id="ccbcc-505">Requirements</span></span>

|<span data-ttu-id="ccbcc-506">要件</span><span class="sxs-lookup"><span data-stu-id="ccbcc-506">Requirement</span></span>| <span data-ttu-id="ccbcc-507">値</span><span class="sxs-lookup"><span data-stu-id="ccbcc-507">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccbcc-508">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccbcc-508">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ccbcc-509">1.5</span><span class="sxs-lookup"><span data-stu-id="ccbcc-509">1.5</span></span> |
|[<span data-ttu-id="ccbcc-510">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccbcc-510">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ccbcc-511">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ccbcc-511">ReadItem</span></span>|
|[<span data-ttu-id="ccbcc-512">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccbcc-512">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ccbcc-513">新規作成と閲覧</span><span class="sxs-lookup"><span data-stu-id="ccbcc-513">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="ccbcc-514">例</span><span class="sxs-lookup"><span data-stu-id="ccbcc-514">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="ccbcc-515">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="ccbcc-515">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="ccbcc-516">Exchange Server から添付ファイルやアイテムを取得するために使用するトークンを含む文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-516">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="ccbcc-p136">`getCallbackTokenAsync` メソッドは、非同期の呼び出しを行なって、ユーザーのメールボックスをホストする Exchange Server から opaque トークンを取得します。コールバック トークンの有効期間は 5 分です。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-p136">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="ccbcc-p137">このトークンと、添付ファイル識別子または項目識別子は、サードパーティーのシステムに渡すことができます。サードパーティーのシステムでは、添付ファイルまたは項目を返すための Exchange Web サービス (EWS) の [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) または [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) 操作を呼び出すのに、トークンをベアラー承認トークンとして使用します。たとえば、リモート サービスを作成して[選択した項目から添付ファイルを取得](/outlook/add-ins/get-attachments-of-an-outlook-item)することができます。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-p137">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="ccbcc-522">アプリでは、閲覧モードで `getCallbackTokenAsync` メソッドを呼び出すために、 **ReadItem** アクセス許可をアプリのマニフェストで指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-522">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="ccbcc-p138">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出して、`getCallbackTokenAsync` メソッドに渡すための項目識別子を取得する必要があります。アプリには、`saveAsync` メソッドを呼び出すために **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-p138">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ccbcc-525">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ccbcc-525">Parameters</span></span>

|<span data-ttu-id="ccbcc-526">名前</span><span class="sxs-lookup"><span data-stu-id="ccbcc-526">Name</span></span>| <span data-ttu-id="ccbcc-527">型</span><span class="sxs-lookup"><span data-stu-id="ccbcc-527">Type</span></span>| <span data-ttu-id="ccbcc-528">属性</span><span class="sxs-lookup"><span data-stu-id="ccbcc-528">Attributes</span></span>| <span data-ttu-id="ccbcc-529">説明</span><span class="sxs-lookup"><span data-stu-id="ccbcc-529">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="ccbcc-530">function</span><span class="sxs-lookup"><span data-stu-id="ccbcc-530">function</span></span>||<span data-ttu-id="ccbcc-p139">メソッドが完了すると、`callback` パラメータに渡された関数が、単一パラメータ `asyncResult` で呼び出されます。このパラメータは、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトです。トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-p139">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="ccbcc-533">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="ccbcc-533">Object</span></span>| <span data-ttu-id="ccbcc-534">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="ccbcc-534">&lt;optional&gt;</span></span>|<span data-ttu-id="ccbcc-535">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-535">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ccbcc-536">要件</span><span class="sxs-lookup"><span data-stu-id="ccbcc-536">Requirements</span></span>

|<span data-ttu-id="ccbcc-537">要件</span><span class="sxs-lookup"><span data-stu-id="ccbcc-537">Requirement</span></span>| <span data-ttu-id="ccbcc-538">値</span><span class="sxs-lookup"><span data-stu-id="ccbcc-538">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccbcc-539">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccbcc-539">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ccbcc-540">1.3</span><span class="sxs-lookup"><span data-stu-id="ccbcc-540">1.3</span></span>|
|[<span data-ttu-id="ccbcc-541">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccbcc-541">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ccbcc-542">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ccbcc-542">ReadItem</span></span>|
|[<span data-ttu-id="ccbcc-543">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccbcc-543">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ccbcc-544">新規作成と閲覧</span><span class="sxs-lookup"><span data-stu-id="ccbcc-544">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="ccbcc-545">例</span><span class="sxs-lookup"><span data-stu-id="ccbcc-545">Example</span></span>

```javascript
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="ccbcc-546">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="ccbcc-546">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="ccbcc-547">ユーザーと Office アドインを識別するトークンを取得します。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-547">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="ccbcc-548">`getUserIdentityTokenAsync` メソッドは、[アドインとユーザーをサードパーティのシステムで識別して認証](/outlook/add-ins/authentication)するのに使用できるトークンを返します。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-548">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="ccbcc-549">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ccbcc-549">Parameters</span></span>

|<span data-ttu-id="ccbcc-550">名前</span><span class="sxs-lookup"><span data-stu-id="ccbcc-550">Name</span></span>| <span data-ttu-id="ccbcc-551">型</span><span class="sxs-lookup"><span data-stu-id="ccbcc-551">Type</span></span>| <span data-ttu-id="ccbcc-552">属性</span><span class="sxs-lookup"><span data-stu-id="ccbcc-552">Attributes</span></span>| <span data-ttu-id="ccbcc-553">説明</span><span class="sxs-lookup"><span data-stu-id="ccbcc-553">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="ccbcc-554">function</span><span class="sxs-lookup"><span data-stu-id="ccbcc-554">function</span></span>||<span data-ttu-id="ccbcc-555">メソッドが完了すると、`callback` パラメータに渡された関数が、単一パラメータ `asyncResult` で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-555">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="ccbcc-556">トークンは、`asyncResult.value`プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-556">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="ccbcc-557">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="ccbcc-557">Object</span></span>| <span data-ttu-id="ccbcc-558">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="ccbcc-558">&lt;optional&gt;</span></span>|<span data-ttu-id="ccbcc-559">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-559">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ccbcc-560">要件</span><span class="sxs-lookup"><span data-stu-id="ccbcc-560">Requirements</span></span>

|<span data-ttu-id="ccbcc-561">要件</span><span class="sxs-lookup"><span data-stu-id="ccbcc-561">Requirement</span></span>| <span data-ttu-id="ccbcc-562">値</span><span class="sxs-lookup"><span data-stu-id="ccbcc-562">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccbcc-563">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccbcc-563">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ccbcc-564">1.0</span><span class="sxs-lookup"><span data-stu-id="ccbcc-564">1.0</span></span>|
|[<span data-ttu-id="ccbcc-565">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccbcc-565">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ccbcc-566">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ccbcc-566">ReadItem</span></span>|
|[<span data-ttu-id="ccbcc-567">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccbcc-567">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ccbcc-568">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ccbcc-568">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ccbcc-569">例</span><span class="sxs-lookup"><span data-stu-id="ccbcc-569">Example</span></span>

```javascript
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="ccbcc-570">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="ccbcc-570">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="ccbcc-571">ユーザーのメールボックスをホストしている Exchange サーバー上の Exchange Web サービス (EWS) のサービスに対して非同期の要求を行います。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-571">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="ccbcc-572">このメソッドは、次のシナリオではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-572">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="ccbcc-573">Outlook for iOS または Outlook for Android で</span><span class="sxs-lookup"><span data-stu-id="ccbcc-573">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="ccbcc-574">アドインが Gmail のメールボックスにロードされる場合</span><span class="sxs-lookup"><span data-stu-id="ccbcc-574">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="ccbcc-575">これらの場合、アドインではユーザーのメールボックスにアクセスするために、代わりに [REST API を使用する](/outlook/add-ins/use-rest-api)必要があります。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-575">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="ccbcc-p140">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange. See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span><span class="sxs-lookup"><span data-stu-id="ccbcc-p140">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange. See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="ccbcc-578">`makeEwsRequestAsync` メソッドでは、フォルダー関連アイテムを要求できません。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-578">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="ccbcc-579">XML 要求では、UTF-8 エンコードを指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-579">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="ccbcc-p141">アドインには、`makeEwsRequestAsync` メソッドを使用するために **ReadWriteMailbox** アクセス許可が必要です。**ReadWriteMailbox** アクセス許可と、`makeEwsRequestAsync` メソッドで呼び出すことのできる EWS 操作の使用の詳細については、「[ユーザーのメールボックスへのメール アドイン アクセスのアクセス許可を指定する](/outlook/add-ins/understanding-outlook-add-in-permissions)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-p141">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="ccbcc-582">サーバー管理者は、クライアント アクセス サーバーの EWS ディレクトリで `OAuthAuthentication` を true に設定して、`makeEwsRequestAsync` メソッドで EWS 要求を行うことができるようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-582">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="ccbcc-583">バージョンの相違点</span><span class="sxs-lookup"><span data-stu-id="ccbcc-583">Version differences</span></span>

<span data-ttu-id="ccbcc-584">バージョン 15.0.4535.1004 より前のバージョンの Outlook で実行しているメール アプリで `makeEwsRequestAsync` メソッドを使用する場合には、エンコード値を `ISO-8859-1` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-584">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="ccbcc-p142">メール アプリが Outlook on the web で実行されている場合には、エンコード値を設定する必要はありません。メールボックスを使用してメール アプリが Outlook で実行されているのか、Outlook on the web で実行されているのかを判断する必要があります。mailbox.diagnostics.hostVersion プロパティを使用すれば、どのバージョンの Outlook が実行されているのかがわかります。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-p142">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ccbcc-588">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ccbcc-588">Parameters</span></span>

|<span data-ttu-id="ccbcc-589">名前</span><span class="sxs-lookup"><span data-stu-id="ccbcc-589">Name</span></span>| <span data-ttu-id="ccbcc-590">型</span><span class="sxs-lookup"><span data-stu-id="ccbcc-590">Type</span></span>| <span data-ttu-id="ccbcc-591">属性</span><span class="sxs-lookup"><span data-stu-id="ccbcc-591">Attributes</span></span>| <span data-ttu-id="ccbcc-592">説明</span><span class="sxs-lookup"><span data-stu-id="ccbcc-592">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="ccbcc-593">文字列</span><span class="sxs-lookup"><span data-stu-id="ccbcc-593">String</span></span>||<span data-ttu-id="ccbcc-594">EWS 要求。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-594">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="ccbcc-595">関数</span><span class="sxs-lookup"><span data-stu-id="ccbcc-595">function</span></span>||<span data-ttu-id="ccbcc-596">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-596">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="ccbcc-p143">The XML result of the EWS call is provided as a string in the `asyncResult.value` property. If the result exceeds 1 MB in size, an error message is returned instead.</span><span class="sxs-lookup"><span data-stu-id="ccbcc-p143">The XML result of the EWS call is provided as a string in the `asyncResult.value` property. If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="ccbcc-599">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="ccbcc-599">Object</span></span>| <span data-ttu-id="ccbcc-600">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="ccbcc-600">&lt;optional&gt;</span></span>|<span data-ttu-id="ccbcc-601">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-601">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ccbcc-602">要件</span><span class="sxs-lookup"><span data-stu-id="ccbcc-602">Requirements</span></span>

|<span data-ttu-id="ccbcc-603">要件</span><span class="sxs-lookup"><span data-stu-id="ccbcc-603">Requirement</span></span>| <span data-ttu-id="ccbcc-604">値</span><span class="sxs-lookup"><span data-stu-id="ccbcc-604">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccbcc-605">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccbcc-605">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ccbcc-606">1.0</span><span class="sxs-lookup"><span data-stu-id="ccbcc-606">1.0</span></span>|
|[<span data-ttu-id="ccbcc-607">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccbcc-607">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ccbcc-608">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="ccbcc-608">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="ccbcc-609">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccbcc-609">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ccbcc-610">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ccbcc-610">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ccbcc-611">例</span><span class="sxs-lookup"><span data-stu-id="ccbcc-611">Example</span></span>

<span data-ttu-id="ccbcc-612">次の例は、`makeEwsRequestAsync` を呼び出し、`GetItem` 操作を使って項目の件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-612">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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

####  <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="ccbcc-613">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="ccbcc-613">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="ccbcc-614">サポートされているイベントの種類のイベント ハンドラーを削除します。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-614">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="ccbcc-615">現在、サポートされている`Office.EventType.ItemChanged`イベント`Office.EventType.OfficeThemeChanged`の種類はおよびです。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-615">Currently, the supported event types are `Office.EventType.ItemChanged` and `Office.EventType.OfficeThemeChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ccbcc-616">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ccbcc-616">Parameters</span></span>

| <span data-ttu-id="ccbcc-617">名前</span><span class="sxs-lookup"><span data-stu-id="ccbcc-617">Name</span></span> | <span data-ttu-id="ccbcc-618">型</span><span class="sxs-lookup"><span data-stu-id="ccbcc-618">Type</span></span> | <span data-ttu-id="ccbcc-619">属性</span><span class="sxs-lookup"><span data-stu-id="ccbcc-619">Attributes</span></span> | <span data-ttu-id="ccbcc-620">説明</span><span class="sxs-lookup"><span data-stu-id="ccbcc-620">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="ccbcc-621">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="ccbcc-621">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="ccbcc-622">ハンドラーを取り消すイベント。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-622">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="ccbcc-623">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="ccbcc-623">Object</span></span> | <span data-ttu-id="ccbcc-624">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="ccbcc-624">&lt;optional&gt;</span></span> | <span data-ttu-id="ccbcc-625">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-625">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="ccbcc-626">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="ccbcc-626">Object</span></span> | <span data-ttu-id="ccbcc-627">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="ccbcc-627">&lt;optional&gt;</span></span> | <span data-ttu-id="ccbcc-628">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-628">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="ccbcc-629">function</span><span class="sxs-lookup"><span data-stu-id="ccbcc-629">function</span></span>| <span data-ttu-id="ccbcc-630">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ccbcc-630">&lt;optional&gt;</span></span>|<span data-ttu-id="ccbcc-631">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="ccbcc-631">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ccbcc-632">要件</span><span class="sxs-lookup"><span data-stu-id="ccbcc-632">Requirements</span></span>

|<span data-ttu-id="ccbcc-633">要件</span><span class="sxs-lookup"><span data-stu-id="ccbcc-633">Requirement</span></span>| <span data-ttu-id="ccbcc-634">値</span><span class="sxs-lookup"><span data-stu-id="ccbcc-634">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccbcc-635">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ccbcc-635">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ccbcc-636">1.5</span><span class="sxs-lookup"><span data-stu-id="ccbcc-636">1.5</span></span> |
|[<span data-ttu-id="ccbcc-637">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ccbcc-637">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ccbcc-638">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ccbcc-638">ReadItem</span></span> |
|[<span data-ttu-id="ccbcc-639">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ccbcc-639">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ccbcc-640">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ccbcc-640">Compose or Read</span></span>|
