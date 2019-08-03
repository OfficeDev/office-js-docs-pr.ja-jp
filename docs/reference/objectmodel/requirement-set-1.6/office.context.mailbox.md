---
title: Office. メールボックス要件セット1.6
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 2ddf26fea8c2285bd577a2f6fb6408431016cc59
ms.sourcegitcommit: 3f5d7f4794e3d3c8bc3a79fa05c54157613b9376
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/02/2019
ms.locfileid: "36064670"
---
# <a name="mailbox"></a><span data-ttu-id="c6202-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="c6202-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="c6202-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="c6202-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="c6202-104">Microsoft Outlook の Outlook アドインオブジェクトモデルへのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="c6202-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c6202-105">要件</span><span class="sxs-lookup"><span data-stu-id="c6202-105">Requirements</span></span>

|<span data-ttu-id="c6202-106">要件</span><span class="sxs-lookup"><span data-stu-id="c6202-106">Requirement</span></span>| <span data-ttu-id="c6202-107">値</span><span class="sxs-lookup"><span data-stu-id="c6202-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6202-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c6202-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c6202-109">1.0</span><span class="sxs-lookup"><span data-stu-id="c6202-109">1.0</span></span>|
|[<span data-ttu-id="c6202-110">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c6202-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c6202-111">制限あり</span><span class="sxs-lookup"><span data-stu-id="c6202-111">Restricted</span></span>|
|[<span data-ttu-id="c6202-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c6202-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c6202-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c6202-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="c6202-114">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="c6202-114">Members and methods</span></span>

| <span data-ttu-id="c6202-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="c6202-115">Member</span></span> | <span data-ttu-id="c6202-116">種類</span><span class="sxs-lookup"><span data-stu-id="c6202-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="c6202-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="c6202-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="c6202-118">メンバー</span><span class="sxs-lookup"><span data-stu-id="c6202-118">Member</span></span> |
| [<span data-ttu-id="c6202-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="c6202-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="c6202-120">メンバー</span><span class="sxs-lookup"><span data-stu-id="c6202-120">Member</span></span> |
| [<span data-ttu-id="c6202-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="c6202-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="c6202-122">メソッド</span><span class="sxs-lookup"><span data-stu-id="c6202-122">Method</span></span> |
| [<span data-ttu-id="c6202-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="c6202-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="c6202-124">メソッド</span><span class="sxs-lookup"><span data-stu-id="c6202-124">Method</span></span> |
| [<span data-ttu-id="c6202-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="c6202-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="c6202-126">メソッド</span><span class="sxs-lookup"><span data-stu-id="c6202-126">Method</span></span> |
| [<span data-ttu-id="c6202-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="c6202-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="c6202-128">メソッド</span><span class="sxs-lookup"><span data-stu-id="c6202-128">Method</span></span> |
| [<span data-ttu-id="c6202-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="c6202-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="c6202-130">メソッド</span><span class="sxs-lookup"><span data-stu-id="c6202-130">Method</span></span> |
| [<span data-ttu-id="c6202-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="c6202-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="c6202-132">メソッド</span><span class="sxs-lookup"><span data-stu-id="c6202-132">Method</span></span> |
| [<span data-ttu-id="c6202-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="c6202-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="c6202-134">メソッド</span><span class="sxs-lookup"><span data-stu-id="c6202-134">Method</span></span> |
| [<span data-ttu-id="c6202-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="c6202-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="c6202-136">メソッド</span><span class="sxs-lookup"><span data-stu-id="c6202-136">Method</span></span> |
| [<span data-ttu-id="c6202-137">displayNewMessageForm</span><span class="sxs-lookup"><span data-stu-id="c6202-137">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="c6202-138">メソッド</span><span class="sxs-lookup"><span data-stu-id="c6202-138">Method</span></span> |
| [<span data-ttu-id="c6202-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="c6202-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="c6202-140">メソッド</span><span class="sxs-lookup"><span data-stu-id="c6202-140">Method</span></span> |
| [<span data-ttu-id="c6202-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="c6202-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="c6202-142">メソッド</span><span class="sxs-lookup"><span data-stu-id="c6202-142">Method</span></span> |
| [<span data-ttu-id="c6202-143">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="c6202-143">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="c6202-144">メソッド</span><span class="sxs-lookup"><span data-stu-id="c6202-144">Method</span></span> |
| [<span data-ttu-id="c6202-145">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="c6202-145">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="c6202-146">メソッド</span><span class="sxs-lookup"><span data-stu-id="c6202-146">Method</span></span> |
| [<span data-ttu-id="c6202-147">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="c6202-147">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="c6202-148">メソッド</span><span class="sxs-lookup"><span data-stu-id="c6202-148">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="c6202-149">名前空間</span><span class="sxs-lookup"><span data-stu-id="c6202-149">Namespaces</span></span>

<span data-ttu-id="c6202-150">[diagnostics](Office.context.mailbox.diagnostics.md):Outlook アドインに診断情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="c6202-150">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="c6202-151">[item](Office.context.mailbox.item.md):Outlook アドインのメッセージや予定にアクセスするためのメソッドとプロパティを提供します。</span><span class="sxs-lookup"><span data-stu-id="c6202-151">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="c6202-152">[userProfile](Office.context.mailbox.userProfile.md):Outlook アドインのユーザーに関する情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="c6202-152">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="c6202-153">メンバー</span><span class="sxs-lookup"><span data-stu-id="c6202-153">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="c6202-154">ewsUrl: String</span><span class="sxs-lookup"><span data-stu-id="c6202-154">ewsUrl: String</span></span>

<span data-ttu-id="c6202-155">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account.</span><span class="sxs-lookup"><span data-stu-id="c6202-155">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account.</span></span> <span data-ttu-id="c6202-156">Read mode only.</span><span class="sxs-lookup"><span data-stu-id="c6202-156">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="c6202-157">このメンバーは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c6202-157">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="c6202-p102">
  `ewsUrl\` 値は、リモート サービスで、ユーザーのメールボックスに EWS 呼び出しを行うために使うことができます。たとえば、[選択したアイテムから添付ファイルを取得する](/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="c6202-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="c6202-160">アプリが閲覧モードで `ewsUrl` メンバーを呼び出すには、アプリのマニフェスト内に **ReadItem** アクセス許可が指定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="c6202-160">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="c6202-p103">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出してから、`ewsUrl` メンバーを使用する必要があります。アプリには、`saveAsync` メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="c6202-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="c6202-163">型</span><span class="sxs-lookup"><span data-stu-id="c6202-163">Type</span></span>

*   <span data-ttu-id="c6202-164">String</span><span class="sxs-lookup"><span data-stu-id="c6202-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c6202-165">要件</span><span class="sxs-lookup"><span data-stu-id="c6202-165">Requirements</span></span>

|<span data-ttu-id="c6202-166">要件</span><span class="sxs-lookup"><span data-stu-id="c6202-166">Requirement</span></span>| <span data-ttu-id="c6202-167">値</span><span class="sxs-lookup"><span data-stu-id="c6202-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6202-168">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c6202-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c6202-169">1.0</span><span class="sxs-lookup"><span data-stu-id="c6202-169">1.0</span></span>|
|[<span data-ttu-id="c6202-170">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c6202-170">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c6202-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6202-171">ReadItem</span></span>|
|[<span data-ttu-id="c6202-172">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c6202-172">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c6202-173">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c6202-173">Compose or Read</span></span>|

#### <a name="resturl-string"></a><span data-ttu-id="c6202-174">Office.context.mailbox.resturl が: String</span><span class="sxs-lookup"><span data-stu-id="c6202-174">restUrl: String</span></span>

<span data-ttu-id="c6202-175">この電子メール アカウントの REST エンドポイントの URL を取得します。</span><span class="sxs-lookup"><span data-stu-id="c6202-175">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="c6202-176">`restUrl` 値は、ユーザーのメールボックスに [REST API](/outlook/rest/) 呼び出しを行うために使用できます。</span><span class="sxs-lookup"><span data-stu-id="c6202-176">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="c6202-177">アプリが閲覧モードで `restUrl` メンバーを呼び出すには、アプリのマニフェスト内に **ReadItem** アクセス許可が指定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="c6202-177">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="c6202-p104">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出してから、`restUrl` メンバーを使用する必要があります。アプリには、`saveAsync` メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="c6202-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="c6202-180">型</span><span class="sxs-lookup"><span data-stu-id="c6202-180">Type</span></span>

*   <span data-ttu-id="c6202-181">String</span><span class="sxs-lookup"><span data-stu-id="c6202-181">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c6202-182">要件</span><span class="sxs-lookup"><span data-stu-id="c6202-182">Requirements</span></span>

|<span data-ttu-id="c6202-183">要件</span><span class="sxs-lookup"><span data-stu-id="c6202-183">Requirement</span></span>| <span data-ttu-id="c6202-184">値</span><span class="sxs-lookup"><span data-stu-id="c6202-184">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6202-185">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c6202-185">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c6202-186">1.5</span><span class="sxs-lookup"><span data-stu-id="c6202-186">1.5</span></span> |
|[<span data-ttu-id="c6202-187">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c6202-187">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c6202-188">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6202-188">ReadItem</span></span>|
|[<span data-ttu-id="c6202-189">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c6202-189">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c6202-190">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c6202-190">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="c6202-191">メソッド</span><span class="sxs-lookup"><span data-stu-id="c6202-191">Methods</span></span>

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="c6202-192">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c6202-192">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="c6202-193">サポートされているイベントのイベント ハンドラーを追加します。</span><span class="sxs-lookup"><span data-stu-id="c6202-193">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="c6202-194">現在サポートされている唯一のイベントの種類は `Office.EventType.ItemChanged` です。これはユーザーが新しいアイテムを選択したときに呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c6202-194">Currently the only supported event type is `Office.EventType.ItemChanged`, which is invoked when the user selects a new item.</span></span> <span data-ttu-id="c6202-195">このイベントは、ピン留め可能な作業ウィンドウを実装するアドインで使用され、現在選択されているアイテムに基づいて作業ウィンドウ UI をアドインで更新できるようにします。</span><span class="sxs-lookup"><span data-stu-id="c6202-195">This event is used by add-ins that implement a pinnable task pane, and allows the add-in to refresh the task pane UI based on the currently selected item.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c6202-196">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c6202-196">Parameters</span></span>

| <span data-ttu-id="c6202-197">名前</span><span class="sxs-lookup"><span data-stu-id="c6202-197">Name</span></span> | <span data-ttu-id="c6202-198">種類</span><span class="sxs-lookup"><span data-stu-id="c6202-198">Type</span></span> | <span data-ttu-id="c6202-199">属性</span><span class="sxs-lookup"><span data-stu-id="c6202-199">Attributes</span></span> | <span data-ttu-id="c6202-200">説明</span><span class="sxs-lookup"><span data-stu-id="c6202-200">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="c6202-201">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="c6202-201">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="c6202-202">ハンドラーを呼び出す必要のあるイベント。</span><span class="sxs-lookup"><span data-stu-id="c6202-202">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="c6202-203">Function</span><span class="sxs-lookup"><span data-stu-id="c6202-203">Function</span></span> || <span data-ttu-id="c6202-p106">イベントを処理する関数。関数は、オブジェクト リテラルである単一パラメーターを受け入れる必要があります。パラメーターの `type` プロパティは、`addHandlerAsync` に渡される `eventType` パラメーターと一致します。</span><span class="sxs-lookup"><span data-stu-id="c6202-p106">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="c6202-207">Object</span><span class="sxs-lookup"><span data-stu-id="c6202-207">Object</span></span> | <span data-ttu-id="c6202-208">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="c6202-208">&lt;optional&gt;</span></span> | <span data-ttu-id="c6202-209">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="c6202-209">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="c6202-210">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="c6202-210">Object</span></span> | <span data-ttu-id="c6202-211">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c6202-211">&lt;optional&gt;</span></span> | <span data-ttu-id="c6202-212">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="c6202-212">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="c6202-213">function</span><span class="sxs-lookup"><span data-stu-id="c6202-213">function</span></span>| <span data-ttu-id="c6202-214">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c6202-214">&lt;optional&gt;</span></span>|<span data-ttu-id="c6202-215">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c6202-215">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c6202-216">要件</span><span class="sxs-lookup"><span data-stu-id="c6202-216">Requirements</span></span>

|<span data-ttu-id="c6202-217">要件</span><span class="sxs-lookup"><span data-stu-id="c6202-217">Requirement</span></span>| <span data-ttu-id="c6202-218">値</span><span class="sxs-lookup"><span data-stu-id="c6202-218">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6202-219">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c6202-219">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c6202-220">1.5</span><span class="sxs-lookup"><span data-stu-id="c6202-220">1.5</span></span> |
|[<span data-ttu-id="c6202-221">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c6202-221">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c6202-222">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6202-222">ReadItem</span></span> |
|[<span data-ttu-id="c6202-223">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c6202-223">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c6202-224">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c6202-224">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c6202-225">例</span><span class="sxs-lookup"><span data-stu-id="c6202-225">Example</span></span>

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

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="c6202-226">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="c6202-226">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="c6202-227">REST 形式のアイテム ID を EWS 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="c6202-227">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="c6202-228">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c6202-228">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="c6202-p107">REST API ([Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) や [Microsoft Graph](https://graph.microsoft.io/) など) で取得されたアイテム ID は、Exchange Web サービス (EWS) に使用される形式とは異なる形式を使用します。`convertToEwsId` メソッドは、REST 形式の ID を EWS 用の適切な形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="c6202-p107">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c6202-231">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c6202-231">Parameters</span></span>

|<span data-ttu-id="c6202-232">名前</span><span class="sxs-lookup"><span data-stu-id="c6202-232">Name</span></span>| <span data-ttu-id="c6202-233">型</span><span class="sxs-lookup"><span data-stu-id="c6202-233">Type</span></span>| <span data-ttu-id="c6202-234">説明</span><span class="sxs-lookup"><span data-stu-id="c6202-234">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="c6202-235">String</span><span class="sxs-lookup"><span data-stu-id="c6202-235">String</span></span>|<span data-ttu-id="c6202-236">Outlook REST API 形式のアイテム ID</span><span class="sxs-lookup"><span data-stu-id="c6202-236">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="c6202-237">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="c6202-237">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.6)|<span data-ttu-id="c6202-238">アイテム ID の取得に使用された Outlook REST API のバージョンを示す値。</span><span class="sxs-lookup"><span data-stu-id="c6202-238">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c6202-239">要件</span><span class="sxs-lookup"><span data-stu-id="c6202-239">Requirements</span></span>

|<span data-ttu-id="c6202-240">要件</span><span class="sxs-lookup"><span data-stu-id="c6202-240">Requirement</span></span>| <span data-ttu-id="c6202-241">値</span><span class="sxs-lookup"><span data-stu-id="c6202-241">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6202-242">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c6202-242">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c6202-243">1.3</span><span class="sxs-lookup"><span data-stu-id="c6202-243">1.3</span></span>|
|[<span data-ttu-id="c6202-244">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c6202-244">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c6202-245">制限あり</span><span class="sxs-lookup"><span data-stu-id="c6202-245">Restricted</span></span>|
|[<span data-ttu-id="c6202-246">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c6202-246">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c6202-247">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c6202-247">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c6202-248">戻り値:</span><span class="sxs-lookup"><span data-stu-id="c6202-248">Returns:</span></span>

<span data-ttu-id="c6202-249">型:String</span><span class="sxs-lookup"><span data-stu-id="c6202-249">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="c6202-250">例</span><span class="sxs-lookup"><span data-stu-id="c6202-250">Example</span></span>

```javascript
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-16"></a><span data-ttu-id="c6202-251">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.6)}</span><span class="sxs-lookup"><span data-stu-id="c6202-251">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.6)}</span></span>

<span data-ttu-id="c6202-252">クライアントのローカル時間で時間情報が含まれている辞書を取得します。</span><span class="sxs-lookup"><span data-stu-id="c6202-252">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="c6202-253">デスクトップまたは web 上の Outlook 用メールアプリは、日付と時刻に異なるタイムゾーンを使用できます。</span><span class="sxs-lookup"><span data-stu-id="c6202-253">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times.</span></span> <span data-ttu-id="c6202-254">デスクトップ上の Outlook では、クライアントコンピューターのタイムゾーンが使用されます。Outlook on the web では、Exchange 管理センター (EAC) で設定されているタイムゾーンが使用されます。</span><span class="sxs-lookup"><span data-stu-id="c6202-254">Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC).</span></span> <span data-ttu-id="c6202-255">日付と時刻の値を処理して、ユーザーインターフェイスに表示される値が、ユーザーが期待するタイムゾーンに常に一致するようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="c6202-255">You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="c6202-256">デスクトップクライアント上の Outlook でメールアプリが実行されている`convertToLocalClientTime`場合、このメソッドは、クライアントコンピューターのタイムゾーンに設定された値を持つ dictionary オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c6202-256">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone.</span></span> <span data-ttu-id="c6202-257">メールアプリが web 上の Outlook で実行されている`convertToLocalClientTime`場合、このメソッドは、EAC で指定されたタイムゾーンに設定された値を持つ dictionary オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c6202-257">If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c6202-258">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c6202-258">Parameters</span></span>

|<span data-ttu-id="c6202-259">名前</span><span class="sxs-lookup"><span data-stu-id="c6202-259">Name</span></span>| <span data-ttu-id="c6202-260">種類</span><span class="sxs-lookup"><span data-stu-id="c6202-260">Type</span></span>| <span data-ttu-id="c6202-261">説明</span><span class="sxs-lookup"><span data-stu-id="c6202-261">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="c6202-262">Date</span><span class="sxs-lookup"><span data-stu-id="c6202-262">Date</span></span>|<span data-ttu-id="c6202-263">日付オブジェクト</span><span class="sxs-lookup"><span data-stu-id="c6202-263">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c6202-264">要件</span><span class="sxs-lookup"><span data-stu-id="c6202-264">Requirements</span></span>

|<span data-ttu-id="c6202-265">要件</span><span class="sxs-lookup"><span data-stu-id="c6202-265">Requirement</span></span>| <span data-ttu-id="c6202-266">値</span><span class="sxs-lookup"><span data-stu-id="c6202-266">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6202-267">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c6202-267">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c6202-268">1.0</span><span class="sxs-lookup"><span data-stu-id="c6202-268">1.0</span></span>|
|[<span data-ttu-id="c6202-269">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c6202-269">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c6202-270">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6202-270">ReadItem</span></span>|
|[<span data-ttu-id="c6202-271">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c6202-271">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c6202-272">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c6202-272">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c6202-273">戻り値:</span><span class="sxs-lookup"><span data-stu-id="c6202-273">Returns:</span></span>

<span data-ttu-id="c6202-274">型:[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="c6202-274">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.6)</span></span>

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="c6202-275">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="c6202-275">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="c6202-276">EWS 形式のアイテム ID を REST 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="c6202-276">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="c6202-277">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c6202-277">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="c6202-p110">EWS または `itemId` プロパティで取得されるアイテム ID は、REST API ([Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) や [Microsoft Graph](https://graph.microsoft.io/) など) に使用される形式とは異なる形式を使用します。`convertToRestId` メソッドは、EWS 形式の ID を REST 用の適切な形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="c6202-p110">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c6202-280">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c6202-280">Parameters</span></span>

|<span data-ttu-id="c6202-281">名前</span><span class="sxs-lookup"><span data-stu-id="c6202-281">Name</span></span>| <span data-ttu-id="c6202-282">種類</span><span class="sxs-lookup"><span data-stu-id="c6202-282">Type</span></span>| <span data-ttu-id="c6202-283">説明</span><span class="sxs-lookup"><span data-stu-id="c6202-283">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="c6202-284">String</span><span class="sxs-lookup"><span data-stu-id="c6202-284">String</span></span>|<span data-ttu-id="c6202-285">Exchange Web サービス (EWS) 形式のアイテム ID</span><span class="sxs-lookup"><span data-stu-id="c6202-285">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="c6202-286">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="c6202-286">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.6)|<span data-ttu-id="c6202-287">変換後の ID を使用する Outlook REST API のバージョンを示す値。</span><span class="sxs-lookup"><span data-stu-id="c6202-287">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c6202-288">要件</span><span class="sxs-lookup"><span data-stu-id="c6202-288">Requirements</span></span>

|<span data-ttu-id="c6202-289">要件</span><span class="sxs-lookup"><span data-stu-id="c6202-289">Requirement</span></span>| <span data-ttu-id="c6202-290">値</span><span class="sxs-lookup"><span data-stu-id="c6202-290">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6202-291">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c6202-291">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c6202-292">1.3</span><span class="sxs-lookup"><span data-stu-id="c6202-292">1.3</span></span>|
|[<span data-ttu-id="c6202-293">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c6202-293">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c6202-294">制限あり</span><span class="sxs-lookup"><span data-stu-id="c6202-294">Restricted</span></span>|
|[<span data-ttu-id="c6202-295">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c6202-295">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c6202-296">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c6202-296">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c6202-297">戻り値:</span><span class="sxs-lookup"><span data-stu-id="c6202-297">Returns:</span></span>

<span data-ttu-id="c6202-298">型:String</span><span class="sxs-lookup"><span data-stu-id="c6202-298">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="c6202-299">例</span><span class="sxs-lookup"><span data-stu-id="c6202-299">Example</span></span>

```javascript
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="c6202-300">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="c6202-300">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="c6202-301">時間情報が含まれているディクショナリから日付オブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="c6202-301">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="c6202-302">`convertToUtcClientTime` メソッドは、ローカルの日付と時刻を含むディクショナリを、ローカルの日付と時刻の正しい値を持つ日付オブジェクトに変換します。</span><span class="sxs-lookup"><span data-stu-id="c6202-302">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c6202-303">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c6202-303">Parameters</span></span>

|<span data-ttu-id="c6202-304">名前</span><span class="sxs-lookup"><span data-stu-id="c6202-304">Name</span></span>| <span data-ttu-id="c6202-305">種類</span><span class="sxs-lookup"><span data-stu-id="c6202-305">Type</span></span>| <span data-ttu-id="c6202-306">説明</span><span class="sxs-lookup"><span data-stu-id="c6202-306">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="c6202-307">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="c6202-307">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.6)|<span data-ttu-id="c6202-308">変換するローカル時刻の値。</span><span class="sxs-lookup"><span data-stu-id="c6202-308">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c6202-309">要件</span><span class="sxs-lookup"><span data-stu-id="c6202-309">Requirements</span></span>

|<span data-ttu-id="c6202-310">要件</span><span class="sxs-lookup"><span data-stu-id="c6202-310">Requirement</span></span>| <span data-ttu-id="c6202-311">値</span><span class="sxs-lookup"><span data-stu-id="c6202-311">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6202-312">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c6202-312">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c6202-313">1.0</span><span class="sxs-lookup"><span data-stu-id="c6202-313">1.0</span></span>|
|[<span data-ttu-id="c6202-314">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c6202-314">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c6202-315">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6202-315">ReadItem</span></span>|
|[<span data-ttu-id="c6202-316">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c6202-316">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c6202-317">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c6202-317">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c6202-318">戻り値:</span><span class="sxs-lookup"><span data-stu-id="c6202-318">Returns:</span></span>

<span data-ttu-id="c6202-319">時間が UTC で表現された日付オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="c6202-319">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="c6202-320">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="c6202-320">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="c6202-321">日付</span><span class="sxs-lookup"><span data-stu-id="c6202-321">Date</span></span></dd>

</dl>

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="c6202-322">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="c6202-322">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="c6202-323">既存の予定を表示します。</span><span class="sxs-lookup"><span data-stu-id="c6202-323">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="c6202-324">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c6202-324">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="c6202-325">`displayAppointmentForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存の予定を開きます。</span><span class="sxs-lookup"><span data-stu-id="c6202-325">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="c6202-326">Outlook on the Mac では、このメソッドを使用して、定期的なアイテムの一部ではない単一の予定を表示したり、定期的なアイテムのマスター予定を表示したりすることはできませんが、一連のインスタンスを表示することはできません。</span><span class="sxs-lookup"><span data-stu-id="c6202-326">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series.</span></span> <span data-ttu-id="c6202-327">これは、Mac 上の Outlook では、定期的なアイテムのインスタンスのプロパティ (アイテム ID を含む) にアクセスできないためです。</span><span class="sxs-lookup"><span data-stu-id="c6202-327">This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="c6202-328">Web 上の Outlook では、フォームの本文が 32 KB 以下の文字である場合にのみ、このメソッドは指定されたフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="c6202-328">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="c6202-329">指定のアイテム識別子が既存の予定を表していない場合は、クライアント コンピューターまたはデバイスで空のウィンドウが開き、エラー メッセージは返されません。</span><span class="sxs-lookup"><span data-stu-id="c6202-329">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c6202-330">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c6202-330">Parameters</span></span>

|<span data-ttu-id="c6202-331">名前</span><span class="sxs-lookup"><span data-stu-id="c6202-331">Name</span></span>| <span data-ttu-id="c6202-332">種類</span><span class="sxs-lookup"><span data-stu-id="c6202-332">Type</span></span>| <span data-ttu-id="c6202-333">説明</span><span class="sxs-lookup"><span data-stu-id="c6202-333">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="c6202-334">String</span><span class="sxs-lookup"><span data-stu-id="c6202-334">String</span></span>|<span data-ttu-id="c6202-335">既存の予定の Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="c6202-335">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c6202-336">要件</span><span class="sxs-lookup"><span data-stu-id="c6202-336">Requirements</span></span>

|<span data-ttu-id="c6202-337">要件</span><span class="sxs-lookup"><span data-stu-id="c6202-337">Requirement</span></span>| <span data-ttu-id="c6202-338">値</span><span class="sxs-lookup"><span data-stu-id="c6202-338">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6202-339">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c6202-339">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c6202-340">1.0</span><span class="sxs-lookup"><span data-stu-id="c6202-340">1.0</span></span>|
|[<span data-ttu-id="c6202-341">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c6202-341">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c6202-342">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6202-342">ReadItem</span></span>|
|[<span data-ttu-id="c6202-343">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c6202-343">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c6202-344">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c6202-344">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c6202-345">例</span><span class="sxs-lookup"><span data-stu-id="c6202-345">Example</span></span>

```javascript
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

#### <a name="displaymessageformitemid"></a><span data-ttu-id="c6202-346">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="c6202-346">displayMessageForm(itemId)</span></span>

<span data-ttu-id="c6202-347">既存のメッセージを表示します。</span><span class="sxs-lookup"><span data-stu-id="c6202-347">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="c6202-348">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c6202-348">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="c6202-349">`displayMessageForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存のメッセージを開きます。</span><span class="sxs-lookup"><span data-stu-id="c6202-349">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="c6202-350">Web 上の Outlook では、フォームの本文が 32 KB の文字数以下の場合にのみ、このメソッドは指定されたフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="c6202-350">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="c6202-351">指定のアイテム識別子が既存のメッセージを表していない場合は、クライアント コンピューターにはメッセージは表示されず、エラー メッセージも返されません。</span><span class="sxs-lookup"><span data-stu-id="c6202-351">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="c6202-p112">予定を表す `itemId` を含む `displayMessageForm` を使用しないでください。既存の予定を表示するには、`displayAppointmentForm` メソッドを使用します。新しい予定を作成するフォームを表示するには、`displayNewAppointmentForm` メソッドを使用します。</span><span class="sxs-lookup"><span data-stu-id="c6202-p112">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c6202-354">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c6202-354">Parameters</span></span>

|<span data-ttu-id="c6202-355">名前</span><span class="sxs-lookup"><span data-stu-id="c6202-355">Name</span></span>| <span data-ttu-id="c6202-356">型</span><span class="sxs-lookup"><span data-stu-id="c6202-356">Type</span></span>| <span data-ttu-id="c6202-357">説明</span><span class="sxs-lookup"><span data-stu-id="c6202-357">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="c6202-358">String</span><span class="sxs-lookup"><span data-stu-id="c6202-358">String</span></span>|<span data-ttu-id="c6202-359">既存のメッセージの Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="c6202-359">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c6202-360">要件</span><span class="sxs-lookup"><span data-stu-id="c6202-360">Requirements</span></span>

|<span data-ttu-id="c6202-361">要件</span><span class="sxs-lookup"><span data-stu-id="c6202-361">Requirement</span></span>| <span data-ttu-id="c6202-362">値</span><span class="sxs-lookup"><span data-stu-id="c6202-362">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6202-363">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c6202-363">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c6202-364">1.0</span><span class="sxs-lookup"><span data-stu-id="c6202-364">1.0</span></span>|
|[<span data-ttu-id="c6202-365">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c6202-365">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c6202-366">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6202-366">ReadItem</span></span>|
|[<span data-ttu-id="c6202-367">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c6202-367">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c6202-368">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c6202-368">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c6202-369">例</span><span class="sxs-lookup"><span data-stu-id="c6202-369">Example</span></span>

```javascript
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="c6202-370">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="c6202-370">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="c6202-371">新しい予定を作成するためのフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="c6202-371">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="c6202-372">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c6202-372">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="c6202-p113">`displayNewAppointmentForm` メソッドを使用すると、ユーザーが新しい予定または会議を作成できるフォームが開きます。パラメーターを指定すると、予定のフォーム フィールドにパラメーターの内容が自動的に設定されます。</span><span class="sxs-lookup"><span data-stu-id="c6202-p113">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="c6202-375">Outlook on the web およびモバイルデバイスでは、このメソッドは常に出席者フィールドを含むフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="c6202-375">In Outlook on the web and mobile devices, this method always displays a form with an attendees field.</span></span> <span data-ttu-id="c6202-376">入力引数として出席者を指定しないと、このメソッドにより **[保存]** ボタンのあるフォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="c6202-376">If you do not specify any attendees as input arguments, the method displays a form with a **Save** button.</span></span> <span data-ttu-id="c6202-377">出席者を指定した場合には、フォームにその出席者と **[送信]** ボタンが表示されます。</span><span class="sxs-lookup"><span data-stu-id="c6202-377">If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="c6202-p115">Outlook リッチ クライアントと Outlook RT で、`requiredAttendees`、`optionalAttendees`、または `resources` パラメーターに出席者またはリソースを指定し、このメソッドを実行すると、**[送信]** ボタンがある会議フォームが表示されます。受信者を指定せずにこのメソッドを実行すると、**[保存して閉じる]** ボタンがある予定フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="c6202-p115">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="c6202-380">パラメーターのいずれかが指定のサイズ制限を超える場合、または不明なパラメーター名が指定されている場合は、例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="c6202-380">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c6202-381">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c6202-381">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="c6202-382">すべてのパラメーターは省略可能です。</span><span class="sxs-lookup"><span data-stu-id="c6202-382">All parameters are optional.</span></span>

|<span data-ttu-id="c6202-383">名前</span><span class="sxs-lookup"><span data-stu-id="c6202-383">Name</span></span>| <span data-ttu-id="c6202-384">型</span><span class="sxs-lookup"><span data-stu-id="c6202-384">Type</span></span>| <span data-ttu-id="c6202-385">説明</span><span class="sxs-lookup"><span data-stu-id="c6202-385">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="c6202-386">Object</span><span class="sxs-lookup"><span data-stu-id="c6202-386">Object</span></span> | <span data-ttu-id="c6202-387">新しい予定を記述するパラメーターのディクショナリ。</span><span class="sxs-lookup"><span data-stu-id="c6202-387">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="c6202-388">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span><span class="sxs-lookup"><span data-stu-id="c6202-388">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span></span> | <span data-ttu-id="c6202-p116">予定に必要な各出席者について、メール アドレスを含む文字列の配列、または `EmailAddressDetails` オブジェクトを含む配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="c6202-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="c6202-391">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span><span class="sxs-lookup"><span data-stu-id="c6202-391">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span></span> | <span data-ttu-id="c6202-p117">予定の各任意出席者について、メール アドレスを含む文字列の配列、または `EmailAddressDetails` オブジェクトを含む配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="c6202-p117">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="c6202-394">Date</span><span class="sxs-lookup"><span data-stu-id="c6202-394">Date</span></span> | <span data-ttu-id="c6202-395">予定の開始日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="c6202-395">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="c6202-396">日付</span><span class="sxs-lookup"><span data-stu-id="c6202-396">Date</span></span> | <span data-ttu-id="c6202-397">予定の終了日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="c6202-397">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="c6202-398">String</span><span class="sxs-lookup"><span data-stu-id="c6202-398">String</span></span> | <span data-ttu-id="c6202-p118">予定の場所を含む文字列。文字列は最大 255 文字に制限されます。</span><span class="sxs-lookup"><span data-stu-id="c6202-p118">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="c6202-401">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="c6202-401">Array.&lt;String&gt;</span></span> | <span data-ttu-id="c6202-p119">予定に必要なリソースを含む文字列の配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="c6202-p119">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="c6202-404">String</span><span class="sxs-lookup"><span data-stu-id="c6202-404">String</span></span> | <span data-ttu-id="c6202-p120">予定の件名を含む文字列です。文字列は最大 255 文字に制限されます。</span><span class="sxs-lookup"><span data-stu-id="c6202-p120">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="c6202-407">String</span><span class="sxs-lookup"><span data-stu-id="c6202-407">String</span></span> | <span data-ttu-id="c6202-p121">予定の本文。本文の内容は、最大サイズが 32 KB に制限されます。</span><span class="sxs-lookup"><span data-stu-id="c6202-p121">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="c6202-410">要件</span><span class="sxs-lookup"><span data-stu-id="c6202-410">Requirements</span></span>

|<span data-ttu-id="c6202-411">要件</span><span class="sxs-lookup"><span data-stu-id="c6202-411">Requirement</span></span>| <span data-ttu-id="c6202-412">値</span><span class="sxs-lookup"><span data-stu-id="c6202-412">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6202-413">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c6202-413">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c6202-414">1.0</span><span class="sxs-lookup"><span data-stu-id="c6202-414">1.0</span></span>|
|[<span data-ttu-id="c6202-415">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c6202-415">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c6202-416">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6202-416">ReadItem</span></span>|
|[<span data-ttu-id="c6202-417">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c6202-417">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c6202-418">読み取り</span><span class="sxs-lookup"><span data-stu-id="c6202-418">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c6202-419">例</span><span class="sxs-lookup"><span data-stu-id="c6202-419">Example</span></span>

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

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="c6202-420">displayNewMessageForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="c6202-420">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="c6202-421">新しいメッセージを作成するためのフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="c6202-421">Displays a form for creating a new message.</span></span>

<span data-ttu-id="c6202-p122">The `displayNewMessageForm` method opens a form that enables the user to create a new message. If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span><span class="sxs-lookup"><span data-stu-id="c6202-p122">The `displayNewMessageForm` method opens a form that enables the user to create a new message. If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="c6202-424">パラメーターのいずれかが指定のサイズ制限を超える場合、または不明なパラメーター名が指定されている場合は、例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="c6202-424">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c6202-425">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c6202-425">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="c6202-426">すべてのパラメーターは省略可能です。</span><span class="sxs-lookup"><span data-stu-id="c6202-426">All parameters are optional.</span></span>

|<span data-ttu-id="c6202-427">名前</span><span class="sxs-lookup"><span data-stu-id="c6202-427">Name</span></span>| <span data-ttu-id="c6202-428">型</span><span class="sxs-lookup"><span data-stu-id="c6202-428">Type</span></span>| <span data-ttu-id="c6202-429">説明</span><span class="sxs-lookup"><span data-stu-id="c6202-429">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="c6202-430">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="c6202-430">Object</span></span> | <span data-ttu-id="c6202-431">新しいメッセージを記述するパラメータの辞書。</span><span class="sxs-lookup"><span data-stu-id="c6202-431">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="c6202-432">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span><span class="sxs-lookup"><span data-stu-id="c6202-432">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span></span> | <span data-ttu-id="c6202-p123">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line. The array is limited to a maximum of 100 entries.</span><span class="sxs-lookup"><span data-stu-id="c6202-p123">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="c6202-435">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span><span class="sxs-lookup"><span data-stu-id="c6202-435">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span></span> | <span data-ttu-id="c6202-p124">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line. The array is limited to a maximum of 100 entries.</span><span class="sxs-lookup"><span data-stu-id="c6202-p124">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="c6202-438">配列。&lt;文字列&gt; | 配列。&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span><span class="sxs-lookup"><span data-stu-id="c6202-438">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span></span> | <span data-ttu-id="c6202-p125">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line. The array is limited to a maximum of 100 entries.</span><span class="sxs-lookup"><span data-stu-id="c6202-p125">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="c6202-441">String</span><span class="sxs-lookup"><span data-stu-id="c6202-441">String</span></span> | <span data-ttu-id="c6202-p126">A string containing the subject of the message. The string is limited to a maximum of 255 characters.</span><span class="sxs-lookup"><span data-stu-id="c6202-p126">A string containing the subject of the message. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="c6202-444">String</span><span class="sxs-lookup"><span data-stu-id="c6202-444">String</span></span> | <span data-ttu-id="c6202-p127">The HTML body of the message. The body content is limited to a maximum size of 32 KB.</span><span class="sxs-lookup"><span data-stu-id="c6202-p127">The HTML body of the message. The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="c6202-447">配列。&lt;オブジェクト&gt;</span><span class="sxs-lookup"><span data-stu-id="c6202-447">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="c6202-448">添付ファイルまたは添付アイテムである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="c6202-448">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="c6202-449">String</span><span class="sxs-lookup"><span data-stu-id="c6202-449">String</span></span> | <span data-ttu-id="c6202-p128">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="c6202-p128">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="c6202-452">String</span><span class="sxs-lookup"><span data-stu-id="c6202-452">String</span></span> | <span data-ttu-id="c6202-453">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="c6202-453">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="c6202-454">文字列</span><span class="sxs-lookup"><span data-stu-id="c6202-454">String</span></span> | <span data-ttu-id="c6202-p129">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="c6202-p129">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="c6202-457">ブール値</span><span class="sxs-lookup"><span data-stu-id="c6202-457">Boolean</span></span> | <span data-ttu-id="c6202-p130">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="c6202-p130">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="c6202-460">String</span><span class="sxs-lookup"><span data-stu-id="c6202-460">String</span></span> | <span data-ttu-id="c6202-p131">Only used if `type` is set to `item`. The EWS item id of the existing e-mail you want to attach to the new message. This is a string up to 100 characters.</span><span class="sxs-lookup"><span data-stu-id="c6202-p131">Only used if `type` is set to `item`. The EWS item id of the existing e-mail you want to attach to the new message. This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="c6202-464">要件</span><span class="sxs-lookup"><span data-stu-id="c6202-464">Requirements</span></span>

|<span data-ttu-id="c6202-465">要件</span><span class="sxs-lookup"><span data-stu-id="c6202-465">Requirement</span></span>| <span data-ttu-id="c6202-466">値</span><span class="sxs-lookup"><span data-stu-id="c6202-466">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6202-467">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c6202-467">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c6202-468">1.6</span><span class="sxs-lookup"><span data-stu-id="c6202-468">1.6</span></span> |
|[<span data-ttu-id="c6202-469">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c6202-469">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c6202-470">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6202-470">ReadItem</span></span>|
|[<span data-ttu-id="c6202-471">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c6202-471">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c6202-472">読み取り</span><span class="sxs-lookup"><span data-stu-id="c6202-472">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c6202-473">例</span><span class="sxs-lookup"><span data-stu-id="c6202-473">Example</span></span>

```javascript
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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="c6202-474">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="c6202-474">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="c6202-475">REST API または Exchange Web サービスを呼び出すために使用するトークンを含む文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="c6202-475">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="c6202-p132">`getCallbackTokenAsync` メソッドは、ユーザーのメールボックスをホストする Exchange Server から不透明なトークンを取得する非同期の呼び出しを行います。コールバック トークンの有効期間は 5 分です。</span><span class="sxs-lookup"><span data-stu-id="c6202-p132">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="c6202-478">可能であれば、アドインでは Exchange Web サービスの代わりに REST API を使用することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="c6202-478">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span> 

<span data-ttu-id="c6202-479">**REST トークン**</span><span class="sxs-lookup"><span data-stu-id="c6202-479">**REST Tokens**</span></span>

<span data-ttu-id="c6202-p133">REST トークンが要求された場合 (`options.isRest = true`)、結果トークンは Exchange Web サービスの呼び出しを認証するためには機能しません。アドインがマニフェストで [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) アクセス許可を指定していない限り、トークンの範囲は現在のアイテムとその添付ファイルへの読み取り専用アクセスに制限されます。`ReadWriteMailbox` アクセス許可が指定されている場合は、結果トークンは、メールを送信する機能など、メール、カレンダー、連絡先への読み取り/書き込みアクセスを付与します。</span><span class="sxs-lookup"><span data-stu-id="c6202-p133">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="c6202-483">アドインでは、`restUrl` プロパティを使用して、REST API 呼び出しを行うときに使用する正しい URL を決定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c6202-483">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="c6202-484">**EWS トークン**</span><span class="sxs-lookup"><span data-stu-id="c6202-484">**EWS Tokens**</span></span>

<span data-ttu-id="c6202-p134">EWS トークンが要求された場合 (`options.isRest = false`)、結果トークンは REST API 呼び出しを認証するためには機能しません。トークンの範囲は、現在のアイテムへのアクセスに制限されます。</span><span class="sxs-lookup"><span data-stu-id="c6202-p134">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="c6202-487">アドインでは、`ewsUrl` プロパティを使用して、EWS 呼び出しを行うときに使用する正しい URL を決定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c6202-487">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c6202-488">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c6202-488">Parameters</span></span>

|<span data-ttu-id="c6202-489">名前</span><span class="sxs-lookup"><span data-stu-id="c6202-489">Name</span></span>| <span data-ttu-id="c6202-490">型</span><span class="sxs-lookup"><span data-stu-id="c6202-490">Type</span></span>| <span data-ttu-id="c6202-491">属性</span><span class="sxs-lookup"><span data-stu-id="c6202-491">Attributes</span></span>| <span data-ttu-id="c6202-492">説明</span><span class="sxs-lookup"><span data-stu-id="c6202-492">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="c6202-493">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="c6202-493">Object</span></span> | <span data-ttu-id="c6202-494">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="c6202-494">&lt;optional&gt;</span></span> | <span data-ttu-id="c6202-495">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="c6202-495">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="c6202-496">Boolean</span><span class="sxs-lookup"><span data-stu-id="c6202-496">Boolean</span></span> |  <span data-ttu-id="c6202-497">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c6202-497">&lt;optional&gt;</span></span> | <span data-ttu-id="c6202-p135">提供されたトークンを Outlook REST API または Exchange Web サービスに使用するかどうかを決定します。既定値は、`false` です。</span><span class="sxs-lookup"><span data-stu-id="c6202-p135">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="c6202-500">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="c6202-500">Object</span></span> |  <span data-ttu-id="c6202-501">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c6202-501">&lt;optional&gt;</span></span> | <span data-ttu-id="c6202-502">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="c6202-502">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="c6202-503">function</span><span class="sxs-lookup"><span data-stu-id="c6202-503">function</span></span>||<span data-ttu-id="c6202-p136">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="c6202-p136">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c6202-506">要件</span><span class="sxs-lookup"><span data-stu-id="c6202-506">Requirements</span></span>

|<span data-ttu-id="c6202-507">要件</span><span class="sxs-lookup"><span data-stu-id="c6202-507">Requirement</span></span>| <span data-ttu-id="c6202-508">値</span><span class="sxs-lookup"><span data-stu-id="c6202-508">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6202-509">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c6202-509">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c6202-510">1.5</span><span class="sxs-lookup"><span data-stu-id="c6202-510">1.5</span></span> |
|[<span data-ttu-id="c6202-511">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c6202-511">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c6202-512">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6202-512">ReadItem</span></span>|
|[<span data-ttu-id="c6202-513">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c6202-513">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c6202-514">新規作成と閲覧</span><span class="sxs-lookup"><span data-stu-id="c6202-514">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="c6202-515">例</span><span class="sxs-lookup"><span data-stu-id="c6202-515">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="c6202-516">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="c6202-516">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="c6202-517">Exchange Server から添付ファイルやアイテムを取得するために使うトークンを含む文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="c6202-517">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="c6202-p137">`getCallbackTokenAsync` メソッドは、ユーザーのメールボックスをホストする Exchange Server から不透明なトークンを取得する非同期の呼び出しを行います。コールバック トークンの有効期間は 5 分です。</span><span class="sxs-lookup"><span data-stu-id="c6202-p137">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="c6202-p138">トークンと予定の識別子またはアイテムの識別子をサードパーティ システムに渡すことができます。サードパーティ システムは、トークンをベアラー承認トークンとして使用し、Exchange Web サービス (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) または [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) 操作を呼び出して、添付ファイルまたはアイテムを返します。たとえば、[選択したアイテムから添付ファイルを取得する](/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="c6202-p138">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="c6202-523">アプリが閲覧モードで `getCallbackTokenAsync` メソッドを呼び出すには、アプリのマニフェスト内に **ReadItem** アクセス許可が指定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="c6202-523">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="c6202-p139">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出してアイテムの識別子を `getCallbackTokenAsync` メソッドに渡す必要があります。アプリには、`saveAsync` メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="c6202-p139">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c6202-526">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c6202-526">Parameters</span></span>

|<span data-ttu-id="c6202-527">名前</span><span class="sxs-lookup"><span data-stu-id="c6202-527">Name</span></span>| <span data-ttu-id="c6202-528">型</span><span class="sxs-lookup"><span data-stu-id="c6202-528">Type</span></span>| <span data-ttu-id="c6202-529">属性</span><span class="sxs-lookup"><span data-stu-id="c6202-529">Attributes</span></span>| <span data-ttu-id="c6202-530">説明</span><span class="sxs-lookup"><span data-stu-id="c6202-530">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="c6202-531">function</span><span class="sxs-lookup"><span data-stu-id="c6202-531">function</span></span>||<span data-ttu-id="c6202-p140">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="c6202-p140">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="c6202-534">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="c6202-534">Object</span></span>| <span data-ttu-id="c6202-535">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c6202-535">&lt;optional&gt;</span></span>|<span data-ttu-id="c6202-536">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="c6202-536">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c6202-537">要件</span><span class="sxs-lookup"><span data-stu-id="c6202-537">Requirements</span></span>

|<span data-ttu-id="c6202-538">要件</span><span class="sxs-lookup"><span data-stu-id="c6202-538">Requirement</span></span>| <span data-ttu-id="c6202-539">値</span><span class="sxs-lookup"><span data-stu-id="c6202-539">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6202-540">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c6202-540">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c6202-541">1.3</span><span class="sxs-lookup"><span data-stu-id="c6202-541">1.3</span></span>|
|[<span data-ttu-id="c6202-542">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c6202-542">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c6202-543">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6202-543">ReadItem</span></span>|
|[<span data-ttu-id="c6202-544">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c6202-544">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c6202-545">新規作成と閲覧</span><span class="sxs-lookup"><span data-stu-id="c6202-545">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="c6202-546">例</span><span class="sxs-lookup"><span data-stu-id="c6202-546">Example</span></span>

```javascript
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="c6202-547">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="c6202-547">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="c6202-548">ユーザーと Office アドインを識別するトークンを取得します。</span><span class="sxs-lookup"><span data-stu-id="c6202-548">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="c6202-549">`getUserIdentityTokenAsync` メソッドは、[アドインとユーザーをサード パーティのシステムで識別して認証](/outlook/add-ins/authentication)することのできるトークンを返します。</span><span class="sxs-lookup"><span data-stu-id="c6202-549">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="c6202-550">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c6202-550">Parameters</span></span>

|<span data-ttu-id="c6202-551">名前</span><span class="sxs-lookup"><span data-stu-id="c6202-551">Name</span></span>| <span data-ttu-id="c6202-552">型</span><span class="sxs-lookup"><span data-stu-id="c6202-552">Type</span></span>| <span data-ttu-id="c6202-553">属性</span><span class="sxs-lookup"><span data-stu-id="c6202-553">Attributes</span></span>| <span data-ttu-id="c6202-554">説明</span><span class="sxs-lookup"><span data-stu-id="c6202-554">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="c6202-555">関数</span><span class="sxs-lookup"><span data-stu-id="c6202-555">function</span></span>||<span data-ttu-id="c6202-556">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c6202-556">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="c6202-557">トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="c6202-557">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="c6202-558">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="c6202-558">Object</span></span>| <span data-ttu-id="c6202-559">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c6202-559">&lt;optional&gt;</span></span>|<span data-ttu-id="c6202-560">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="c6202-560">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c6202-561">要件</span><span class="sxs-lookup"><span data-stu-id="c6202-561">Requirements</span></span>

|<span data-ttu-id="c6202-562">要件</span><span class="sxs-lookup"><span data-stu-id="c6202-562">Requirement</span></span>| <span data-ttu-id="c6202-563">値</span><span class="sxs-lookup"><span data-stu-id="c6202-563">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6202-564">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c6202-564">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c6202-565">1.0</span><span class="sxs-lookup"><span data-stu-id="c6202-565">1.0</span></span>|
|[<span data-ttu-id="c6202-566">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c6202-566">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c6202-567">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6202-567">ReadItem</span></span>|
|[<span data-ttu-id="c6202-568">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c6202-568">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c6202-569">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c6202-569">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c6202-570">例</span><span class="sxs-lookup"><span data-stu-id="c6202-570">Example</span></span>

```javascript
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="c6202-571">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="c6202-571">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="c6202-572">ユーザーのメールボックスをホストしている Exchange サーバー上の Exchange Web サービス (EWS) のサービスに対して非同期の要求を行います。</span><span class="sxs-lookup"><span data-stu-id="c6202-572">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="c6202-573">このメソッドは、次のシナリオではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c6202-573">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="c6202-574">Outlook on iOS または Android</span><span class="sxs-lookup"><span data-stu-id="c6202-574">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="c6202-575">アドインが Gmail のメールボックスに読み込まれる場合</span><span class="sxs-lookup"><span data-stu-id="c6202-575">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="c6202-576">このような場合は、アドインでは [REST API を使用](/outlook/add-ins/use-rest-api)して、代わりにユーザーのメールボックスにアクセスする必要があります。</span><span class="sxs-lookup"><span data-stu-id="c6202-576">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="c6202-p141">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange. See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span><span class="sxs-lookup"><span data-stu-id="c6202-p141">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange. See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="c6202-579">`makeEwsRequestAsync` メソッドでは、フォルダー関連アイテムを要求できません。</span><span class="sxs-lookup"><span data-stu-id="c6202-579">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="c6202-580">XML 要求では UTF-8 エンコードを指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c6202-580">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="c6202-p142">`makeEwsRequestAsync` メソッドを使用するには、アドインに **ReadWriteMailbox** アクセス許可が必要です。**ReadWriteMailbox** アクセス許可と、`makeEwsRequestAsync` メソッドで呼び出せる EWS 操作の使い方については、「[ユーザーのメールボックスへのメール アドイン アクセスのアクセス許可を指定する](/outlook/add-ins/understanding-outlook-add-in-permissions)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c6202-p142">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="c6202-583">サーバー管理者は、クライアント アクセス サーバーの EWS ディレクトリで `OAuthAuthentication` を true に設定して、`makeEwsRequestAsync` メソッドで EWS 要求を行うことができるようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="c6202-583">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="c6202-584">バージョンの相違点</span><span class="sxs-lookup"><span data-stu-id="c6202-584">Version differences</span></span>

<span data-ttu-id="c6202-585">バージョン 15.0.4535.1004 より前のバージョンの Outlook で実行しているメール アプリで `makeEwsRequestAsync` メソッドを使う場合は、エンコード値を `ISO-8859-1` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c6202-585">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="c6202-p143">Outlook on the web でメール アプリを実行している場合は、エンコード値を設定する必要はありません。mailbox.diagnostics.hostName プロパティを使って、メール アプリを Outlook で実行しているのか、Outlook on the web で実行しているのかを確認できます。mailbox.diagnostics.hostVersion プロパティを使って、どのバージョンの Outlook を使って実行しているかを確認できます。</span><span class="sxs-lookup"><span data-stu-id="c6202-p143">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c6202-589">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c6202-589">Parameters</span></span>

|<span data-ttu-id="c6202-590">名前</span><span class="sxs-lookup"><span data-stu-id="c6202-590">Name</span></span>| <span data-ttu-id="c6202-591">型</span><span class="sxs-lookup"><span data-stu-id="c6202-591">Type</span></span>| <span data-ttu-id="c6202-592">属性</span><span class="sxs-lookup"><span data-stu-id="c6202-592">Attributes</span></span>| <span data-ttu-id="c6202-593">説明</span><span class="sxs-lookup"><span data-stu-id="c6202-593">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="c6202-594">String</span><span class="sxs-lookup"><span data-stu-id="c6202-594">String</span></span>||<span data-ttu-id="c6202-595">EWS 要求です。</span><span class="sxs-lookup"><span data-stu-id="c6202-595">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="c6202-596">関数</span><span class="sxs-lookup"><span data-stu-id="c6202-596">function</span></span>||<span data-ttu-id="c6202-597">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c6202-597">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="c6202-p144">The XML result of the EWS call is provided as a string in the `asyncResult.value` property. If the result exceeds 1 MB in size, an error message is returned instead.</span><span class="sxs-lookup"><span data-stu-id="c6202-p144">The XML result of the EWS call is provided as a string in the `asyncResult.value` property. If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="c6202-600">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="c6202-600">Object</span></span>| <span data-ttu-id="c6202-601">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c6202-601">&lt;optional&gt;</span></span>|<span data-ttu-id="c6202-602">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="c6202-602">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c6202-603">要件</span><span class="sxs-lookup"><span data-stu-id="c6202-603">Requirements</span></span>

|<span data-ttu-id="c6202-604">要件</span><span class="sxs-lookup"><span data-stu-id="c6202-604">Requirement</span></span>| <span data-ttu-id="c6202-605">値</span><span class="sxs-lookup"><span data-stu-id="c6202-605">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6202-606">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c6202-606">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c6202-607">1.0</span><span class="sxs-lookup"><span data-stu-id="c6202-607">1.0</span></span>|
|[<span data-ttu-id="c6202-608">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c6202-608">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c6202-609">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="c6202-609">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="c6202-610">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c6202-610">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c6202-611">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c6202-611">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c6202-612">例</span><span class="sxs-lookup"><span data-stu-id="c6202-612">Example</span></span>

<span data-ttu-id="c6202-613">次の例は、`GetItem` 操作を使ってアイテムの件名を取得するため、`makeEwsRequestAsync` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="c6202-613">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="c6202-614">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c6202-614">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="c6202-615">サポートされているイベントの種類のイベント ハンドラーを削除します。</span><span class="sxs-lookup"><span data-stu-id="c6202-615">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="c6202-616">現在、サポートされているイベントの種類は `Office.EventType.ItemChanged` だけです。</span><span class="sxs-lookup"><span data-stu-id="c6202-616">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c6202-617">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c6202-617">Parameters</span></span>

| <span data-ttu-id="c6202-618">名前</span><span class="sxs-lookup"><span data-stu-id="c6202-618">Name</span></span> | <span data-ttu-id="c6202-619">型</span><span class="sxs-lookup"><span data-stu-id="c6202-619">Type</span></span> | <span data-ttu-id="c6202-620">属性</span><span class="sxs-lookup"><span data-stu-id="c6202-620">Attributes</span></span> | <span data-ttu-id="c6202-621">説明</span><span class="sxs-lookup"><span data-stu-id="c6202-621">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="c6202-622">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="c6202-622">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="c6202-623">ハンドラーを取り消すイベント。</span><span class="sxs-lookup"><span data-stu-id="c6202-623">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="c6202-624">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="c6202-624">Object</span></span> | <span data-ttu-id="c6202-625">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="c6202-625">&lt;optional&gt;</span></span> | <span data-ttu-id="c6202-626">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="c6202-626">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="c6202-627">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="c6202-627">Object</span></span> | <span data-ttu-id="c6202-628">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c6202-628">&lt;optional&gt;</span></span> | <span data-ttu-id="c6202-629">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="c6202-629">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="c6202-630">function</span><span class="sxs-lookup"><span data-stu-id="c6202-630">function</span></span>| <span data-ttu-id="c6202-631">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c6202-631">&lt;optional&gt;</span></span>|<span data-ttu-id="c6202-632">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c6202-632">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c6202-633">要件</span><span class="sxs-lookup"><span data-stu-id="c6202-633">Requirements</span></span>

|<span data-ttu-id="c6202-634">要件</span><span class="sxs-lookup"><span data-stu-id="c6202-634">Requirement</span></span>| <span data-ttu-id="c6202-635">値</span><span class="sxs-lookup"><span data-stu-id="c6202-635">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6202-636">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c6202-636">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c6202-637">1.5</span><span class="sxs-lookup"><span data-stu-id="c6202-637">1.5</span></span> |
|[<span data-ttu-id="c6202-638">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c6202-638">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c6202-639">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c6202-639">ReadItem</span></span> |
|[<span data-ttu-id="c6202-640">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c6202-640">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c6202-641">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c6202-641">Compose or Read</span></span>|
