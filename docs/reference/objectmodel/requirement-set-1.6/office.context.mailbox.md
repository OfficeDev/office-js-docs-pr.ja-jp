---
title: Office. メールボックス要件セット1.6
description: ''
ms.date: 08/30/2019
localization_priority: Normal
ms.openlocfilehash: 82a7039602c1896488e6a2358cf345bc157b79de
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/30/2019
ms.locfileid: "36695981"
---
# <a name="mailbox"></a><span data-ttu-id="b18d8-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="b18d8-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="b18d8-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="b18d8-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="b18d8-104">Microsoft Outlook の Outlook アドインオブジェクトモデルへのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="b18d8-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="b18d8-105">要件</span><span class="sxs-lookup"><span data-stu-id="b18d8-105">Requirements</span></span>

|<span data-ttu-id="b18d8-106">要件</span><span class="sxs-lookup"><span data-stu-id="b18d8-106">Requirement</span></span>| <span data-ttu-id="b18d8-107">値</span><span class="sxs-lookup"><span data-stu-id="b18d8-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="b18d8-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b18d8-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b18d8-109">1.0</span><span class="sxs-lookup"><span data-stu-id="b18d8-109">1.0</span></span>|
|[<span data-ttu-id="b18d8-110">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b18d8-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b18d8-111">制限あり</span><span class="sxs-lookup"><span data-stu-id="b18d8-111">Restricted</span></span>|
|[<span data-ttu-id="b18d8-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b18d8-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b18d8-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="b18d8-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="b18d8-114">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="b18d8-114">Members and methods</span></span>

| <span data-ttu-id="b18d8-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="b18d8-115">Member</span></span> | <span data-ttu-id="b18d8-116">種類</span><span class="sxs-lookup"><span data-stu-id="b18d8-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="b18d8-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="b18d8-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="b18d8-118">メンバー</span><span class="sxs-lookup"><span data-stu-id="b18d8-118">Member</span></span> |
| [<span data-ttu-id="b18d8-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="b18d8-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="b18d8-120">メンバー</span><span class="sxs-lookup"><span data-stu-id="b18d8-120">Member</span></span> |
| [<span data-ttu-id="b18d8-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="b18d8-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="b18d8-122">メソッド</span><span class="sxs-lookup"><span data-stu-id="b18d8-122">Method</span></span> |
| [<span data-ttu-id="b18d8-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="b18d8-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="b18d8-124">メソッド</span><span class="sxs-lookup"><span data-stu-id="b18d8-124">Method</span></span> |
| [<span data-ttu-id="b18d8-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="b18d8-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="b18d8-126">メソッド</span><span class="sxs-lookup"><span data-stu-id="b18d8-126">Method</span></span> |
| [<span data-ttu-id="b18d8-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="b18d8-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="b18d8-128">メソッド</span><span class="sxs-lookup"><span data-stu-id="b18d8-128">Method</span></span> |
| [<span data-ttu-id="b18d8-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="b18d8-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="b18d8-130">メソッド</span><span class="sxs-lookup"><span data-stu-id="b18d8-130">Method</span></span> |
| [<span data-ttu-id="b18d8-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="b18d8-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="b18d8-132">メソッド</span><span class="sxs-lookup"><span data-stu-id="b18d8-132">Method</span></span> |
| [<span data-ttu-id="b18d8-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="b18d8-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="b18d8-134">メソッド</span><span class="sxs-lookup"><span data-stu-id="b18d8-134">Method</span></span> |
| [<span data-ttu-id="b18d8-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="b18d8-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="b18d8-136">メソッド</span><span class="sxs-lookup"><span data-stu-id="b18d8-136">Method</span></span> |
| [<span data-ttu-id="b18d8-137">displayNewMessageForm</span><span class="sxs-lookup"><span data-stu-id="b18d8-137">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="b18d8-138">メソッド</span><span class="sxs-lookup"><span data-stu-id="b18d8-138">Method</span></span> |
| [<span data-ttu-id="b18d8-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="b18d8-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="b18d8-140">メソッド</span><span class="sxs-lookup"><span data-stu-id="b18d8-140">Method</span></span> |
| [<span data-ttu-id="b18d8-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="b18d8-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="b18d8-142">メソッド</span><span class="sxs-lookup"><span data-stu-id="b18d8-142">Method</span></span> |
| [<span data-ttu-id="b18d8-143">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="b18d8-143">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="b18d8-144">メソッド</span><span class="sxs-lookup"><span data-stu-id="b18d8-144">Method</span></span> |
| [<span data-ttu-id="b18d8-145">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="b18d8-145">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="b18d8-146">メソッド</span><span class="sxs-lookup"><span data-stu-id="b18d8-146">Method</span></span> |
| [<span data-ttu-id="b18d8-147">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="b18d8-147">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="b18d8-148">メソッド</span><span class="sxs-lookup"><span data-stu-id="b18d8-148">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="b18d8-149">名前空間</span><span class="sxs-lookup"><span data-stu-id="b18d8-149">Namespaces</span></span>

<span data-ttu-id="b18d8-150">[diagnostics](Office.context.mailbox.diagnostics.md):Outlook アドインに診断情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="b18d8-150">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="b18d8-151">[item](Office.context.mailbox.item.md):Outlook アドインのメッセージや予定にアクセスするためのメソッドとプロパティを提供します。</span><span class="sxs-lookup"><span data-stu-id="b18d8-151">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="b18d8-152">[userProfile](Office.context.mailbox.userProfile.md):Outlook アドインのユーザーに関する情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="b18d8-152">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="b18d8-153">メンバー</span><span class="sxs-lookup"><span data-stu-id="b18d8-153">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="b18d8-154">ewsUrl: String</span><span class="sxs-lookup"><span data-stu-id="b18d8-154">ewsUrl: String</span></span>

<span data-ttu-id="b18d8-155">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account.</span><span class="sxs-lookup"><span data-stu-id="b18d8-155">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account.</span></span> <span data-ttu-id="b18d8-156">Read mode only.</span><span class="sxs-lookup"><span data-stu-id="b18d8-156">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="b18d8-157">このメンバーは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="b18d8-157">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="b18d8-p102">
  `ewsUrl\` 値は、リモート サービスで、ユーザーのメールボックスに EWS 呼び出しを行うために使うことができます。たとえば、[選択したアイテムから添付ファイルを取得する](/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="b18d8-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="b18d8-160">アプリが閲覧モードで `ewsUrl` メンバーを呼び出すには、アプリのマニフェスト内に **ReadItem** アクセス許可が指定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="b18d8-160">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="b18d8-p103">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出してから、`ewsUrl` メンバーを使用する必要があります。アプリには、`saveAsync` メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="b18d8-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="b18d8-163">型</span><span class="sxs-lookup"><span data-stu-id="b18d8-163">Type</span></span>

*   <span data-ttu-id="b18d8-164">String</span><span class="sxs-lookup"><span data-stu-id="b18d8-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b18d8-165">要件</span><span class="sxs-lookup"><span data-stu-id="b18d8-165">Requirements</span></span>

|<span data-ttu-id="b18d8-166">要件</span><span class="sxs-lookup"><span data-stu-id="b18d8-166">Requirement</span></span>| <span data-ttu-id="b18d8-167">値</span><span class="sxs-lookup"><span data-stu-id="b18d8-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="b18d8-168">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b18d8-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b18d8-169">1.0</span><span class="sxs-lookup"><span data-stu-id="b18d8-169">1.0</span></span>|
|[<span data-ttu-id="b18d8-170">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b18d8-170">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b18d8-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b18d8-171">ReadItem</span></span>|
|[<span data-ttu-id="b18d8-172">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b18d8-172">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b18d8-173">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="b18d8-173">Compose or Read</span></span>|

<br>

---
---

#### <a name="resturl-string"></a><span data-ttu-id="b18d8-174">Office.context.mailbox.resturl が: String</span><span class="sxs-lookup"><span data-stu-id="b18d8-174">restUrl: String</span></span>

<span data-ttu-id="b18d8-175">この電子メール アカウントの REST エンドポイントの URL を取得します。</span><span class="sxs-lookup"><span data-stu-id="b18d8-175">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="b18d8-176">`restUrl` 値は、ユーザーのメールボックスに [REST API](/outlook/rest/) 呼び出しを行うために使用できます。</span><span class="sxs-lookup"><span data-stu-id="b18d8-176">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="b18d8-177">アプリが閲覧モードで `restUrl` メンバーを呼び出すには、アプリのマニフェスト内に **ReadItem** アクセス許可が指定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="b18d8-177">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="b18d8-p104">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出してから、`restUrl` メンバーを使用する必要があります。アプリには、`saveAsync` メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="b18d8-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="b18d8-180">型</span><span class="sxs-lookup"><span data-stu-id="b18d8-180">Type</span></span>

*   <span data-ttu-id="b18d8-181">String</span><span class="sxs-lookup"><span data-stu-id="b18d8-181">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b18d8-182">要件</span><span class="sxs-lookup"><span data-stu-id="b18d8-182">Requirements</span></span>

|<span data-ttu-id="b18d8-183">要件</span><span class="sxs-lookup"><span data-stu-id="b18d8-183">Requirement</span></span>| <span data-ttu-id="b18d8-184">値</span><span class="sxs-lookup"><span data-stu-id="b18d8-184">Value</span></span>|
|---|---|
|[<span data-ttu-id="b18d8-185">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b18d8-185">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b18d8-186">1.5</span><span class="sxs-lookup"><span data-stu-id="b18d8-186">1.5</span></span> |
|[<span data-ttu-id="b18d8-187">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b18d8-187">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b18d8-188">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b18d8-188">ReadItem</span></span>|
|[<span data-ttu-id="b18d8-189">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b18d8-189">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b18d8-190">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="b18d8-190">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="b18d8-191">メソッド</span><span class="sxs-lookup"><span data-stu-id="b18d8-191">Methods</span></span>

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="b18d8-192">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="b18d8-192">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="b18d8-193">サポートされているイベントのイベント ハンドラーを追加します。</span><span class="sxs-lookup"><span data-stu-id="b18d8-193">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="b18d8-194">現在サポートされている唯一のイベントの種類は `Office.EventType.ItemChanged` です。これはユーザーが新しいアイテムを選択したときに呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="b18d8-194">Currently the only supported event type is `Office.EventType.ItemChanged`, which is invoked when the user selects a new item.</span></span> <span data-ttu-id="b18d8-195">このイベントは、ピン留め可能な作業ウィンドウを実装するアドインで使用され、現在選択されているアイテムに基づいて作業ウィンドウ UI をアドインで更新できるようにします。</span><span class="sxs-lookup"><span data-stu-id="b18d8-195">This event is used by add-ins that implement a pinnable task pane, and allows the add-in to refresh the task pane UI based on the currently selected item.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b18d8-196">パラメーター</span><span class="sxs-lookup"><span data-stu-id="b18d8-196">Parameters</span></span>

| <span data-ttu-id="b18d8-197">名前</span><span class="sxs-lookup"><span data-stu-id="b18d8-197">Name</span></span> | <span data-ttu-id="b18d8-198">種類</span><span class="sxs-lookup"><span data-stu-id="b18d8-198">Type</span></span> | <span data-ttu-id="b18d8-199">属性</span><span class="sxs-lookup"><span data-stu-id="b18d8-199">Attributes</span></span> | <span data-ttu-id="b18d8-200">説明</span><span class="sxs-lookup"><span data-stu-id="b18d8-200">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="b18d8-201">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="b18d8-201">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="b18d8-202">ハンドラーを呼び出す必要のあるイベント。</span><span class="sxs-lookup"><span data-stu-id="b18d8-202">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="b18d8-203">Function</span><span class="sxs-lookup"><span data-stu-id="b18d8-203">Function</span></span> || <span data-ttu-id="b18d8-p106">イベントを処理する関数。関数は、オブジェクト リテラルである単一パラメーターを受け入れる必要があります。パラメーターの `type` プロパティは、`addHandlerAsync` に渡される `eventType` パラメーターと一致します。</span><span class="sxs-lookup"><span data-stu-id="b18d8-p106">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="b18d8-207">Object</span><span class="sxs-lookup"><span data-stu-id="b18d8-207">Object</span></span> | <span data-ttu-id="b18d8-208">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="b18d8-208">&lt;optional&gt;</span></span> | <span data-ttu-id="b18d8-209">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="b18d8-209">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="b18d8-210">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="b18d8-210">Object</span></span> | <span data-ttu-id="b18d8-211">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="b18d8-211">&lt;optional&gt;</span></span> | <span data-ttu-id="b18d8-212">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="b18d8-212">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="b18d8-213">function</span><span class="sxs-lookup"><span data-stu-id="b18d8-213">function</span></span>| <span data-ttu-id="b18d8-214">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b18d8-214">&lt;optional&gt;</span></span>|<span data-ttu-id="b18d8-215">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="b18d8-215">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b18d8-216">要件</span><span class="sxs-lookup"><span data-stu-id="b18d8-216">Requirements</span></span>

|<span data-ttu-id="b18d8-217">要件</span><span class="sxs-lookup"><span data-stu-id="b18d8-217">Requirement</span></span>| <span data-ttu-id="b18d8-218">値</span><span class="sxs-lookup"><span data-stu-id="b18d8-218">Value</span></span>|
|---|---|
|[<span data-ttu-id="b18d8-219">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b18d8-219">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b18d8-220">1.5</span><span class="sxs-lookup"><span data-stu-id="b18d8-220">1.5</span></span> |
|[<span data-ttu-id="b18d8-221">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b18d8-221">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b18d8-222">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b18d8-222">ReadItem</span></span> |
|[<span data-ttu-id="b18d8-223">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b18d8-223">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b18d8-224">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="b18d8-224">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b18d8-225">例</span><span class="sxs-lookup"><span data-stu-id="b18d8-225">Example</span></span>

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

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="b18d8-226">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="b18d8-226">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="b18d8-227">REST 形式のアイテム ID を EWS 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="b18d8-227">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="b18d8-228">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="b18d8-228">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="b18d8-p107">REST API ([Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) や [Microsoft Graph](https://graph.microsoft.io/) など) で取得されたアイテム ID は、Exchange Web サービス (EWS) に使用される形式とは異なる形式を使用します。`convertToEwsId` メソッドは、REST 形式の ID を EWS 用の適切な形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="b18d8-p107">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b18d8-231">パラメーター</span><span class="sxs-lookup"><span data-stu-id="b18d8-231">Parameters</span></span>

|<span data-ttu-id="b18d8-232">名前</span><span class="sxs-lookup"><span data-stu-id="b18d8-232">Name</span></span>| <span data-ttu-id="b18d8-233">型</span><span class="sxs-lookup"><span data-stu-id="b18d8-233">Type</span></span>| <span data-ttu-id="b18d8-234">説明</span><span class="sxs-lookup"><span data-stu-id="b18d8-234">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="b18d8-235">String</span><span class="sxs-lookup"><span data-stu-id="b18d8-235">String</span></span>|<span data-ttu-id="b18d8-236">Outlook REST API 形式のアイテム ID</span><span class="sxs-lookup"><span data-stu-id="b18d8-236">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="b18d8-237">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="b18d8-237">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.6)|<span data-ttu-id="b18d8-238">アイテム ID の取得に使用された Outlook REST API のバージョンを示す値。</span><span class="sxs-lookup"><span data-stu-id="b18d8-238">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b18d8-239">要件</span><span class="sxs-lookup"><span data-stu-id="b18d8-239">Requirements</span></span>

|<span data-ttu-id="b18d8-240">要件</span><span class="sxs-lookup"><span data-stu-id="b18d8-240">Requirement</span></span>| <span data-ttu-id="b18d8-241">値</span><span class="sxs-lookup"><span data-stu-id="b18d8-241">Value</span></span>|
|---|---|
|[<span data-ttu-id="b18d8-242">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b18d8-242">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b18d8-243">1.3</span><span class="sxs-lookup"><span data-stu-id="b18d8-243">1.3</span></span>|
|[<span data-ttu-id="b18d8-244">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b18d8-244">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b18d8-245">制限あり</span><span class="sxs-lookup"><span data-stu-id="b18d8-245">Restricted</span></span>|
|[<span data-ttu-id="b18d8-246">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b18d8-246">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b18d8-247">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="b18d8-247">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b18d8-248">戻り値:</span><span class="sxs-lookup"><span data-stu-id="b18d8-248">Returns:</span></span>

<span data-ttu-id="b18d8-249">型:String</span><span class="sxs-lookup"><span data-stu-id="b18d8-249">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="b18d8-250">例</span><span class="sxs-lookup"><span data-stu-id="b18d8-250">Example</span></span>

```js
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-16"></a><span data-ttu-id="b18d8-251">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.6)}</span><span class="sxs-lookup"><span data-stu-id="b18d8-251">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.6)}</span></span>

<span data-ttu-id="b18d8-252">クライアントのローカル時間で時間情報が含まれている辞書を取得します。</span><span class="sxs-lookup"><span data-stu-id="b18d8-252">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="b18d8-253">デスクトップまたは web 上の Outlook 用メールアプリは、日付と時刻に異なるタイムゾーンを使用できます。</span><span class="sxs-lookup"><span data-stu-id="b18d8-253">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times.</span></span> <span data-ttu-id="b18d8-254">デスクトップ上の Outlook では、クライアントコンピューターのタイムゾーンが使用されます。Outlook on the web では、Exchange 管理センター (EAC) で設定されているタイムゾーンが使用されます。</span><span class="sxs-lookup"><span data-stu-id="b18d8-254">Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC).</span></span> <span data-ttu-id="b18d8-255">日付と時刻の値を処理して、ユーザーインターフェイスに表示される値が、ユーザーが期待するタイムゾーンに常に一致するようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="b18d8-255">You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="b18d8-256">デスクトップクライアント上の Outlook でメールアプリが実行されている`convertToLocalClientTime`場合、このメソッドは、クライアントコンピューターのタイムゾーンに設定された値を持つ dictionary オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="b18d8-256">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone.</span></span> <span data-ttu-id="b18d8-257">メールアプリが web 上の Outlook で実行されている`convertToLocalClientTime`場合、このメソッドは、EAC で指定されたタイムゾーンに設定された値を持つ dictionary オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="b18d8-257">If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b18d8-258">パラメーター</span><span class="sxs-lookup"><span data-stu-id="b18d8-258">Parameters</span></span>

|<span data-ttu-id="b18d8-259">名前</span><span class="sxs-lookup"><span data-stu-id="b18d8-259">Name</span></span>| <span data-ttu-id="b18d8-260">型</span><span class="sxs-lookup"><span data-stu-id="b18d8-260">Type</span></span>| <span data-ttu-id="b18d8-261">説明</span><span class="sxs-lookup"><span data-stu-id="b18d8-261">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="b18d8-262">日付</span><span class="sxs-lookup"><span data-stu-id="b18d8-262">Date</span></span>|<span data-ttu-id="b18d8-263">日付オブジェクト</span><span class="sxs-lookup"><span data-stu-id="b18d8-263">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b18d8-264">要件</span><span class="sxs-lookup"><span data-stu-id="b18d8-264">Requirements</span></span>

|<span data-ttu-id="b18d8-265">要件</span><span class="sxs-lookup"><span data-stu-id="b18d8-265">Requirement</span></span>| <span data-ttu-id="b18d8-266">値</span><span class="sxs-lookup"><span data-stu-id="b18d8-266">Value</span></span>|
|---|---|
|[<span data-ttu-id="b18d8-267">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b18d8-267">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b18d8-268">1.0</span><span class="sxs-lookup"><span data-stu-id="b18d8-268">1.0</span></span>|
|[<span data-ttu-id="b18d8-269">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b18d8-269">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b18d8-270">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b18d8-270">ReadItem</span></span>|
|[<span data-ttu-id="b18d8-271">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b18d8-271">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b18d8-272">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="b18d8-272">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b18d8-273">戻り値:</span><span class="sxs-lookup"><span data-stu-id="b18d8-273">Returns:</span></span>

<span data-ttu-id="b18d8-274">型:[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="b18d8-274">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.6)</span></span>

<br>

---
---

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="b18d8-275">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="b18d8-275">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="b18d8-276">EWS 形式のアイテム ID を REST 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="b18d8-276">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="b18d8-277">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="b18d8-277">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="b18d8-p110">EWS または `itemId` プロパティで取得されるアイテム ID は、REST API ([Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) や [Microsoft Graph](https://graph.microsoft.io/) など) に使用される形式とは異なる形式を使用します。`convertToRestId` メソッドは、EWS 形式の ID を REST 用の適切な形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="b18d8-p110">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b18d8-280">パラメーター</span><span class="sxs-lookup"><span data-stu-id="b18d8-280">Parameters</span></span>

|<span data-ttu-id="b18d8-281">名前</span><span class="sxs-lookup"><span data-stu-id="b18d8-281">Name</span></span>| <span data-ttu-id="b18d8-282">型</span><span class="sxs-lookup"><span data-stu-id="b18d8-282">Type</span></span>| <span data-ttu-id="b18d8-283">説明</span><span class="sxs-lookup"><span data-stu-id="b18d8-283">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="b18d8-284">String</span><span class="sxs-lookup"><span data-stu-id="b18d8-284">String</span></span>|<span data-ttu-id="b18d8-285">Exchange Web サービス (EWS) 形式のアイテム ID</span><span class="sxs-lookup"><span data-stu-id="b18d8-285">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="b18d8-286">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="b18d8-286">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.6)|<span data-ttu-id="b18d8-287">変換後の ID を使用する Outlook REST API のバージョンを示す値。</span><span class="sxs-lookup"><span data-stu-id="b18d8-287">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b18d8-288">要件</span><span class="sxs-lookup"><span data-stu-id="b18d8-288">Requirements</span></span>

|<span data-ttu-id="b18d8-289">要件</span><span class="sxs-lookup"><span data-stu-id="b18d8-289">Requirement</span></span>| <span data-ttu-id="b18d8-290">値</span><span class="sxs-lookup"><span data-stu-id="b18d8-290">Value</span></span>|
|---|---|
|[<span data-ttu-id="b18d8-291">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b18d8-291">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b18d8-292">1.3</span><span class="sxs-lookup"><span data-stu-id="b18d8-292">1.3</span></span>|
|[<span data-ttu-id="b18d8-293">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b18d8-293">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b18d8-294">制限あり</span><span class="sxs-lookup"><span data-stu-id="b18d8-294">Restricted</span></span>|
|[<span data-ttu-id="b18d8-295">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b18d8-295">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b18d8-296">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="b18d8-296">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b18d8-297">戻り値:</span><span class="sxs-lookup"><span data-stu-id="b18d8-297">Returns:</span></span>

<span data-ttu-id="b18d8-298">型:String</span><span class="sxs-lookup"><span data-stu-id="b18d8-298">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="b18d8-299">例</span><span class="sxs-lookup"><span data-stu-id="b18d8-299">Example</span></span>

```js
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="b18d8-300">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="b18d8-300">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="b18d8-301">時間情報が含まれているディクショナリから日付オブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="b18d8-301">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="b18d8-302">`convertToUtcClientTime` メソッドは、ローカルの日付と時刻を含むディクショナリを、ローカルの日付と時刻の正しい値を持つ日付オブジェクトに変換します。</span><span class="sxs-lookup"><span data-stu-id="b18d8-302">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b18d8-303">パラメーター</span><span class="sxs-lookup"><span data-stu-id="b18d8-303">Parameters</span></span>

|<span data-ttu-id="b18d8-304">名前</span><span class="sxs-lookup"><span data-stu-id="b18d8-304">Name</span></span>| <span data-ttu-id="b18d8-305">型</span><span class="sxs-lookup"><span data-stu-id="b18d8-305">Type</span></span>| <span data-ttu-id="b18d8-306">説明</span><span class="sxs-lookup"><span data-stu-id="b18d8-306">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="b18d8-307">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="b18d8-307">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.6)|<span data-ttu-id="b18d8-308">変換するローカル時刻の値。</span><span class="sxs-lookup"><span data-stu-id="b18d8-308">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b18d8-309">要件</span><span class="sxs-lookup"><span data-stu-id="b18d8-309">Requirements</span></span>

|<span data-ttu-id="b18d8-310">要件</span><span class="sxs-lookup"><span data-stu-id="b18d8-310">Requirement</span></span>| <span data-ttu-id="b18d8-311">値</span><span class="sxs-lookup"><span data-stu-id="b18d8-311">Value</span></span>|
|---|---|
|[<span data-ttu-id="b18d8-312">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b18d8-312">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b18d8-313">1.0</span><span class="sxs-lookup"><span data-stu-id="b18d8-313">1.0</span></span>|
|[<span data-ttu-id="b18d8-314">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b18d8-314">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b18d8-315">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b18d8-315">ReadItem</span></span>|
|[<span data-ttu-id="b18d8-316">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b18d8-316">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b18d8-317">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="b18d8-317">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b18d8-318">戻り値:</span><span class="sxs-lookup"><span data-stu-id="b18d8-318">Returns:</span></span>

<span data-ttu-id="b18d8-319">時間が UTC で表現された日付オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="b18d8-319">A Date object with the time expressed in UTC.</span></span>

<span data-ttu-id="b18d8-320">型: Date</span><span class="sxs-lookup"><span data-stu-id="b18d8-320">Type: Date</span></span>

##### <a name="example"></a><span data-ttu-id="b18d8-321">例</span><span class="sxs-lookup"><span data-stu-id="b18d8-321">Example</span></span>

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

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="b18d8-322">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="b18d8-322">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="b18d8-323">既存の予定を表示します。</span><span class="sxs-lookup"><span data-stu-id="b18d8-323">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="b18d8-324">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="b18d8-324">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="b18d8-325">`displayAppointmentForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存の予定を開きます。</span><span class="sxs-lookup"><span data-stu-id="b18d8-325">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="b18d8-326">Outlook on the Mac では、このメソッドを使用して、定期的なアイテムの一部ではない単一の予定を表示したり、定期的なアイテムのマスター予定を表示したりすることはできませんが、一連のインスタンスを表示することはできません。</span><span class="sxs-lookup"><span data-stu-id="b18d8-326">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series.</span></span> <span data-ttu-id="b18d8-327">これは、Mac 上の Outlook では、定期的なアイテムのインスタンスのプロパティ (アイテム ID を含む) にアクセスできないためです。</span><span class="sxs-lookup"><span data-stu-id="b18d8-327">This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="b18d8-328">Web 上の Outlook では、フォームの本文が 32 KB 以下の文字である場合にのみ、このメソッドは指定されたフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="b18d8-328">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="b18d8-329">指定のアイテム識別子が既存の予定を表していない場合は、クライアント コンピューターまたはデバイスで空のウィンドウが開き、エラー メッセージは返されません。</span><span class="sxs-lookup"><span data-stu-id="b18d8-329">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b18d8-330">パラメーター</span><span class="sxs-lookup"><span data-stu-id="b18d8-330">Parameters</span></span>

|<span data-ttu-id="b18d8-331">名前</span><span class="sxs-lookup"><span data-stu-id="b18d8-331">Name</span></span>| <span data-ttu-id="b18d8-332">型</span><span class="sxs-lookup"><span data-stu-id="b18d8-332">Type</span></span>| <span data-ttu-id="b18d8-333">説明</span><span class="sxs-lookup"><span data-stu-id="b18d8-333">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="b18d8-334">String</span><span class="sxs-lookup"><span data-stu-id="b18d8-334">String</span></span>|<span data-ttu-id="b18d8-335">既存の予定の Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="b18d8-335">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b18d8-336">要件</span><span class="sxs-lookup"><span data-stu-id="b18d8-336">Requirements</span></span>

|<span data-ttu-id="b18d8-337">要件</span><span class="sxs-lookup"><span data-stu-id="b18d8-337">Requirement</span></span>| <span data-ttu-id="b18d8-338">値</span><span class="sxs-lookup"><span data-stu-id="b18d8-338">Value</span></span>|
|---|---|
|[<span data-ttu-id="b18d8-339">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b18d8-339">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b18d8-340">1.0</span><span class="sxs-lookup"><span data-stu-id="b18d8-340">1.0</span></span>|
|[<span data-ttu-id="b18d8-341">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b18d8-341">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b18d8-342">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b18d8-342">ReadItem</span></span>|
|[<span data-ttu-id="b18d8-343">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b18d8-343">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b18d8-344">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="b18d8-344">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b18d8-345">例</span><span class="sxs-lookup"><span data-stu-id="b18d8-345">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

<br>

---
---

#### <a name="displaymessageformitemid"></a><span data-ttu-id="b18d8-346">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="b18d8-346">displayMessageForm(itemId)</span></span>

<span data-ttu-id="b18d8-347">既存のメッセージを表示します。</span><span class="sxs-lookup"><span data-stu-id="b18d8-347">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="b18d8-348">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="b18d8-348">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="b18d8-349">`displayMessageForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存のメッセージを開きます。</span><span class="sxs-lookup"><span data-stu-id="b18d8-349">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="b18d8-350">Web 上の Outlook では、フォームの本文が 32 KB の文字数以下の場合にのみ、このメソッドは指定されたフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="b18d8-350">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="b18d8-351">指定のアイテム識別子が既存のメッセージを表していない場合は、クライアント コンピューターにはメッセージは表示されず、エラー メッセージも返されません。</span><span class="sxs-lookup"><span data-stu-id="b18d8-351">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="b18d8-p112">予定を表す `itemId` を含む `displayMessageForm` を使用しないでください。既存の予定を表示するには、`displayAppointmentForm` メソッドを使用します。新しい予定を作成するフォームを表示するには、`displayNewAppointmentForm` メソッドを使用します。</span><span class="sxs-lookup"><span data-stu-id="b18d8-p112">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b18d8-354">パラメーター</span><span class="sxs-lookup"><span data-stu-id="b18d8-354">Parameters</span></span>

|<span data-ttu-id="b18d8-355">名前</span><span class="sxs-lookup"><span data-stu-id="b18d8-355">Name</span></span>| <span data-ttu-id="b18d8-356">型</span><span class="sxs-lookup"><span data-stu-id="b18d8-356">Type</span></span>| <span data-ttu-id="b18d8-357">説明</span><span class="sxs-lookup"><span data-stu-id="b18d8-357">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="b18d8-358">String</span><span class="sxs-lookup"><span data-stu-id="b18d8-358">String</span></span>|<span data-ttu-id="b18d8-359">既存のメッセージの Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="b18d8-359">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b18d8-360">要件</span><span class="sxs-lookup"><span data-stu-id="b18d8-360">Requirements</span></span>

|<span data-ttu-id="b18d8-361">要件</span><span class="sxs-lookup"><span data-stu-id="b18d8-361">Requirement</span></span>| <span data-ttu-id="b18d8-362">値</span><span class="sxs-lookup"><span data-stu-id="b18d8-362">Value</span></span>|
|---|---|
|[<span data-ttu-id="b18d8-363">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b18d8-363">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b18d8-364">1.0</span><span class="sxs-lookup"><span data-stu-id="b18d8-364">1.0</span></span>|
|[<span data-ttu-id="b18d8-365">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b18d8-365">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b18d8-366">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b18d8-366">ReadItem</span></span>|
|[<span data-ttu-id="b18d8-367">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b18d8-367">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b18d8-368">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="b18d8-368">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b18d8-369">例</span><span class="sxs-lookup"><span data-stu-id="b18d8-369">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

<br>

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="b18d8-370">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="b18d8-370">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="b18d8-371">新しい予定を作成するためのフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="b18d8-371">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="b18d8-372">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="b18d8-372">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="b18d8-p113">`displayNewAppointmentForm` メソッドを使用すると、ユーザーが新しい予定または会議を作成できるフォームが開きます。パラメーターを指定すると、予定のフォーム フィールドにパラメーターの内容が自動的に設定されます。</span><span class="sxs-lookup"><span data-stu-id="b18d8-p113">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="b18d8-375">Outlook on the web およびモバイルデバイスでは、このメソッドは常に出席者フィールドを含むフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="b18d8-375">In Outlook on the web and mobile devices, this method always displays a form with an attendees field.</span></span> <span data-ttu-id="b18d8-376">入力引数として出席者を指定しないと、このメソッドにより **[保存]** ボタンのあるフォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="b18d8-376">If you do not specify any attendees as input arguments, the method displays a form with a **Save** button.</span></span> <span data-ttu-id="b18d8-377">出席者を指定した場合には、フォームにその出席者と **[送信]** ボタンが表示されます。</span><span class="sxs-lookup"><span data-stu-id="b18d8-377">If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="b18d8-p115">Outlook リッチ クライアントと Outlook RT で、`requiredAttendees`、`optionalAttendees`、または `resources` パラメーターに出席者またはリソースを指定し、このメソッドを実行すると、**[送信]** ボタンがある会議フォームが表示されます。受信者を指定せずにこのメソッドを実行すると、**[保存して閉じる]** ボタンがある予定フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="b18d8-p115">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="b18d8-380">パラメーターのいずれかが指定のサイズ制限を超える場合、または不明なパラメーター名が指定されている場合は、例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="b18d8-380">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b18d8-381">パラメーター</span><span class="sxs-lookup"><span data-stu-id="b18d8-381">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="b18d8-382">すべてのパラメーターは省略可能です。</span><span class="sxs-lookup"><span data-stu-id="b18d8-382">All parameters are optional.</span></span>

|<span data-ttu-id="b18d8-383">名前</span><span class="sxs-lookup"><span data-stu-id="b18d8-383">Name</span></span>| <span data-ttu-id="b18d8-384">型</span><span class="sxs-lookup"><span data-stu-id="b18d8-384">Type</span></span>| <span data-ttu-id="b18d8-385">説明</span><span class="sxs-lookup"><span data-stu-id="b18d8-385">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="b18d8-386">Object</span><span class="sxs-lookup"><span data-stu-id="b18d8-386">Object</span></span> | <span data-ttu-id="b18d8-387">新しい予定を記述するパラメーターのディクショナリ。</span><span class="sxs-lookup"><span data-stu-id="b18d8-387">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="b18d8-388">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span><span class="sxs-lookup"><span data-stu-id="b18d8-388">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span></span> | <span data-ttu-id="b18d8-p116">予定に必要な各出席者について、メール アドレスを含む文字列の配列、または `EmailAddressDetails` オブジェクトを含む配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="b18d8-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="b18d8-391">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span><span class="sxs-lookup"><span data-stu-id="b18d8-391">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span></span> | <span data-ttu-id="b18d8-p117">予定の各任意出席者について、メール アドレスを含む文字列の配列、または `EmailAddressDetails` オブジェクトを含む配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="b18d8-p117">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="b18d8-394">Date</span><span class="sxs-lookup"><span data-stu-id="b18d8-394">Date</span></span> | <span data-ttu-id="b18d8-395">予定の開始日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="b18d8-395">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="b18d8-396">日付</span><span class="sxs-lookup"><span data-stu-id="b18d8-396">Date</span></span> | <span data-ttu-id="b18d8-397">予定の終了日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="b18d8-397">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="b18d8-398">String</span><span class="sxs-lookup"><span data-stu-id="b18d8-398">String</span></span> | <span data-ttu-id="b18d8-p118">予定の場所を含む文字列。文字列は最大 255 文字に制限されます。</span><span class="sxs-lookup"><span data-stu-id="b18d8-p118">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="b18d8-401">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="b18d8-401">Array.&lt;String&gt;</span></span> | <span data-ttu-id="b18d8-p119">予定に必要なリソースを含む文字列の配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="b18d8-p119">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="b18d8-404">String</span><span class="sxs-lookup"><span data-stu-id="b18d8-404">String</span></span> | <span data-ttu-id="b18d8-p120">予定の件名を含む文字列です。文字列は最大 255 文字に制限されます。</span><span class="sxs-lookup"><span data-stu-id="b18d8-p120">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="b18d8-407">String</span><span class="sxs-lookup"><span data-stu-id="b18d8-407">String</span></span> | <span data-ttu-id="b18d8-p121">予定の本文。本文の内容は、最大サイズが 32 KB に制限されます。</span><span class="sxs-lookup"><span data-stu-id="b18d8-p121">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="b18d8-410">要件</span><span class="sxs-lookup"><span data-stu-id="b18d8-410">Requirements</span></span>

|<span data-ttu-id="b18d8-411">要件</span><span class="sxs-lookup"><span data-stu-id="b18d8-411">Requirement</span></span>| <span data-ttu-id="b18d8-412">値</span><span class="sxs-lookup"><span data-stu-id="b18d8-412">Value</span></span>|
|---|---|
|[<span data-ttu-id="b18d8-413">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b18d8-413">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b18d8-414">1.0</span><span class="sxs-lookup"><span data-stu-id="b18d8-414">1.0</span></span>|
|[<span data-ttu-id="b18d8-415">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b18d8-415">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b18d8-416">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b18d8-416">ReadItem</span></span>|
|[<span data-ttu-id="b18d8-417">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b18d8-417">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b18d8-418">読み取り</span><span class="sxs-lookup"><span data-stu-id="b18d8-418">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b18d8-419">例</span><span class="sxs-lookup"><span data-stu-id="b18d8-419">Example</span></span>

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

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="b18d8-420">displayNewMessageForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="b18d8-420">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="b18d8-421">新しいメッセージを作成するためのフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="b18d8-421">Displays a form for creating a new message.</span></span>

<span data-ttu-id="b18d8-p122">The `displayNewMessageForm` method opens a form that enables the user to create a new message. If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span><span class="sxs-lookup"><span data-stu-id="b18d8-p122">The `displayNewMessageForm` method opens a form that enables the user to create a new message. If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="b18d8-424">パラメーターのいずれかが指定のサイズ制限を超える場合、または不明なパラメーター名が指定されている場合は、例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="b18d8-424">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b18d8-425">パラメーター</span><span class="sxs-lookup"><span data-stu-id="b18d8-425">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="b18d8-426">すべてのパラメーターは省略可能です。</span><span class="sxs-lookup"><span data-stu-id="b18d8-426">All parameters are optional.</span></span>

|<span data-ttu-id="b18d8-427">名前</span><span class="sxs-lookup"><span data-stu-id="b18d8-427">Name</span></span>| <span data-ttu-id="b18d8-428">型</span><span class="sxs-lookup"><span data-stu-id="b18d8-428">Type</span></span>| <span data-ttu-id="b18d8-429">説明</span><span class="sxs-lookup"><span data-stu-id="b18d8-429">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="b18d8-430">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="b18d8-430">Object</span></span> | <span data-ttu-id="b18d8-431">新しいメッセージを記述するパラメータの辞書。</span><span class="sxs-lookup"><span data-stu-id="b18d8-431">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="b18d8-432">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span><span class="sxs-lookup"><span data-stu-id="b18d8-432">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span></span> | <span data-ttu-id="b18d8-p123">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line. The array is limited to a maximum of 100 entries.</span><span class="sxs-lookup"><span data-stu-id="b18d8-p123">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="b18d8-435">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span><span class="sxs-lookup"><span data-stu-id="b18d8-435">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span></span> | <span data-ttu-id="b18d8-p124">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line. The array is limited to a maximum of 100 entries.</span><span class="sxs-lookup"><span data-stu-id="b18d8-p124">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="b18d8-438">配列。&lt;文字列&gt; | 配列。&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span><span class="sxs-lookup"><span data-stu-id="b18d8-438">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span></span> | <span data-ttu-id="b18d8-p125">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line. The array is limited to a maximum of 100 entries.</span><span class="sxs-lookup"><span data-stu-id="b18d8-p125">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="b18d8-441">String</span><span class="sxs-lookup"><span data-stu-id="b18d8-441">String</span></span> | <span data-ttu-id="b18d8-p126">A string containing the subject of the message. The string is limited to a maximum of 255 characters.</span><span class="sxs-lookup"><span data-stu-id="b18d8-p126">A string containing the subject of the message. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="b18d8-444">String</span><span class="sxs-lookup"><span data-stu-id="b18d8-444">String</span></span> | <span data-ttu-id="b18d8-p127">The HTML body of the message. The body content is limited to a maximum size of 32 KB.</span><span class="sxs-lookup"><span data-stu-id="b18d8-p127">The HTML body of the message. The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="b18d8-447">配列。&lt;オブジェクト&gt;</span><span class="sxs-lookup"><span data-stu-id="b18d8-447">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="b18d8-448">添付ファイルまたは添付アイテムである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="b18d8-448">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="b18d8-449">String</span><span class="sxs-lookup"><span data-stu-id="b18d8-449">String</span></span> | <span data-ttu-id="b18d8-p128">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="b18d8-p128">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="b18d8-452">String</span><span class="sxs-lookup"><span data-stu-id="b18d8-452">String</span></span> | <span data-ttu-id="b18d8-453">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="b18d8-453">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="b18d8-454">文字列</span><span class="sxs-lookup"><span data-stu-id="b18d8-454">String</span></span> | <span data-ttu-id="b18d8-p129">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="b18d8-p129">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="b18d8-457">ブール値</span><span class="sxs-lookup"><span data-stu-id="b18d8-457">Boolean</span></span> | <span data-ttu-id="b18d8-p130">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="b18d8-p130">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="b18d8-460">String</span><span class="sxs-lookup"><span data-stu-id="b18d8-460">String</span></span> | <span data-ttu-id="b18d8-p131">Only used if `type` is set to `item`. The EWS item id of the existing e-mail you want to attach to the new message. This is a string up to 100 characters.</span><span class="sxs-lookup"><span data-stu-id="b18d8-p131">Only used if `type` is set to `item`. The EWS item id of the existing e-mail you want to attach to the new message. This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="b18d8-464">要件</span><span class="sxs-lookup"><span data-stu-id="b18d8-464">Requirements</span></span>

|<span data-ttu-id="b18d8-465">要件</span><span class="sxs-lookup"><span data-stu-id="b18d8-465">Requirement</span></span>| <span data-ttu-id="b18d8-466">値</span><span class="sxs-lookup"><span data-stu-id="b18d8-466">Value</span></span>|
|---|---|
|[<span data-ttu-id="b18d8-467">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b18d8-467">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b18d8-468">1.6</span><span class="sxs-lookup"><span data-stu-id="b18d8-468">1.6</span></span> |
|[<span data-ttu-id="b18d8-469">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b18d8-469">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b18d8-470">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b18d8-470">ReadItem</span></span>|
|[<span data-ttu-id="b18d8-471">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b18d8-471">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b18d8-472">読み取り</span><span class="sxs-lookup"><span data-stu-id="b18d8-472">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b18d8-473">例</span><span class="sxs-lookup"><span data-stu-id="b18d8-473">Example</span></span>

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

<br>

---
---

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="b18d8-474">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="b18d8-474">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="b18d8-475">REST API または Exchange Web サービスを呼び出すために使用するトークンを含む文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="b18d8-475">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="b18d8-p132">`getCallbackTokenAsync` メソッドは、ユーザーのメールボックスをホストする Exchange Server から不透明なトークンを取得する非同期の呼び出しを行います。コールバック トークンの有効期間は 5 分です。</span><span class="sxs-lookup"><span data-stu-id="b18d8-p132">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="b18d8-478">可能であれば、アドインでは Exchange Web サービスの代わりに REST API を使用することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="b18d8-478">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span> 

<span data-ttu-id="b18d8-479">**REST トークン**</span><span class="sxs-lookup"><span data-stu-id="b18d8-479">**REST Tokens**</span></span>

<span data-ttu-id="b18d8-p133">REST トークンが要求された場合 (`options.isRest = true`)、結果トークンは Exchange Web サービスの呼び出しを認証するためには機能しません。アドインがマニフェストで [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) アクセス許可を指定していない限り、トークンの範囲は現在のアイテムとその添付ファイルへの読み取り専用アクセスに制限されます。`ReadWriteMailbox` アクセス許可が指定されている場合は、結果トークンは、メールを送信する機能など、メール、カレンダー、連絡先への読み取り/書き込みアクセスを付与します。</span><span class="sxs-lookup"><span data-stu-id="b18d8-p133">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="b18d8-483">アドインでは、`restUrl` プロパティを使用して、REST API 呼び出しを行うときに使用する正しい URL を決定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="b18d8-483">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="b18d8-484">**EWS トークン**</span><span class="sxs-lookup"><span data-stu-id="b18d8-484">**EWS Tokens**</span></span>

<span data-ttu-id="b18d8-p134">EWS トークンが要求された場合 (`options.isRest = false`)、結果トークンは REST API 呼び出しを認証するためには機能しません。トークンの範囲は、現在のアイテムへのアクセスに制限されます。</span><span class="sxs-lookup"><span data-stu-id="b18d8-p134">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="b18d8-487">アドインでは、`ewsUrl` プロパティを使用して、EWS 呼び出しを行うときに使用する正しい URL を決定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="b18d8-487">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b18d8-488">パラメーター</span><span class="sxs-lookup"><span data-stu-id="b18d8-488">Parameters</span></span>

|<span data-ttu-id="b18d8-489">名前</span><span class="sxs-lookup"><span data-stu-id="b18d8-489">Name</span></span>| <span data-ttu-id="b18d8-490">型</span><span class="sxs-lookup"><span data-stu-id="b18d8-490">Type</span></span>| <span data-ttu-id="b18d8-491">属性</span><span class="sxs-lookup"><span data-stu-id="b18d8-491">Attributes</span></span>| <span data-ttu-id="b18d8-492">説明</span><span class="sxs-lookup"><span data-stu-id="b18d8-492">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="b18d8-493">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="b18d8-493">Object</span></span> | <span data-ttu-id="b18d8-494">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="b18d8-494">&lt;optional&gt;</span></span> | <span data-ttu-id="b18d8-495">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="b18d8-495">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="b18d8-496">Boolean</span><span class="sxs-lookup"><span data-stu-id="b18d8-496">Boolean</span></span> |  <span data-ttu-id="b18d8-497">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b18d8-497">&lt;optional&gt;</span></span> | <span data-ttu-id="b18d8-p135">提供されたトークンを Outlook REST API または Exchange Web サービスに使用するかどうかを決定します。既定値は、`false` です。</span><span class="sxs-lookup"><span data-stu-id="b18d8-p135">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="b18d8-500">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="b18d8-500">Object</span></span> |  <span data-ttu-id="b18d8-501">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="b18d8-501">&lt;optional&gt;</span></span> | <span data-ttu-id="b18d8-502">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="b18d8-502">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="b18d8-503">function</span><span class="sxs-lookup"><span data-stu-id="b18d8-503">function</span></span>||<span data-ttu-id="b18d8-504">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="b18d8-504">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="b18d8-505">トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="b18d8-505">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="b18d8-506">エラーが発生した場合は`asyncResult.error` 、 `asyncResult.diagnostics`プロパティとプロパティによって追加情報が提供されることがあります。</span><span class="sxs-lookup"><span data-stu-id="b18d8-506">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|

##### <a name="errors"></a><span data-ttu-id="b18d8-507">エラー</span><span class="sxs-lookup"><span data-stu-id="b18d8-507">Errors</span></span>

|<span data-ttu-id="b18d8-508">エラー コード</span><span class="sxs-lookup"><span data-stu-id="b18d8-508">Error code</span></span>|<span data-ttu-id="b18d8-509">説明</span><span class="sxs-lookup"><span data-stu-id="b18d8-509">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="b18d8-510">要求が失敗しました。</span><span class="sxs-lookup"><span data-stu-id="b18d8-510">The request has failed.</span></span> <span data-ttu-id="b18d8-511">HTTP エラーコードについては、diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="b18d8-511">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="b18d8-512">Exchange サーバーがエラーを返しました。</span><span class="sxs-lookup"><span data-stu-id="b18d8-512">The Exchange server returned an error.</span></span> <span data-ttu-id="b18d8-513">詳細については、「diagnostics オブジェクト」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="b18d8-513">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="b18d8-514">ユーザーがネットワークに接続されていません。</span><span class="sxs-lookup"><span data-stu-id="b18d8-514">The user is no longer connected to the network.</span></span> <span data-ttu-id="b18d8-515">ネットワーク接続を確認し、もう一度実行してください。</span><span class="sxs-lookup"><span data-stu-id="b18d8-515">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b18d8-516">要件</span><span class="sxs-lookup"><span data-stu-id="b18d8-516">Requirements</span></span>

|<span data-ttu-id="b18d8-517">要件</span><span class="sxs-lookup"><span data-stu-id="b18d8-517">Requirement</span></span>| <span data-ttu-id="b18d8-518">値</span><span class="sxs-lookup"><span data-stu-id="b18d8-518">Value</span></span>|
|---|---|
|[<span data-ttu-id="b18d8-519">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b18d8-519">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b18d8-520">1.5</span><span class="sxs-lookup"><span data-stu-id="b18d8-520">1.5</span></span> |
|[<span data-ttu-id="b18d8-521">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b18d8-521">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b18d8-522">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b18d8-522">ReadItem</span></span>|
|[<span data-ttu-id="b18d8-523">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b18d8-523">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b18d8-524">新規作成と閲覧</span><span class="sxs-lookup"><span data-stu-id="b18d8-524">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="b18d8-525">例</span><span class="sxs-lookup"><span data-stu-id="b18d8-525">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="b18d8-526">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="b18d8-526">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="b18d8-527">Exchange Server から添付ファイルやアイテムを取得するために使うトークンを含む文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="b18d8-527">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="b18d8-p139">`getCallbackTokenAsync` メソッドは、ユーザーのメールボックスをホストする Exchange Server から不透明なトークンを取得する非同期の呼び出しを行います。コールバック トークンの有効期間は 5 分です。</span><span class="sxs-lookup"><span data-stu-id="b18d8-p139">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="b18d8-p140">トークンと予定の識別子またはアイテムの識別子をサードパーティ システムに渡すことができます。サードパーティ システムは、トークンをベアラー承認トークンとして使用し、Exchange Web サービス (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) または [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) 操作を呼び出して、添付ファイルまたはアイテムを返します。たとえば、[選択したアイテムから添付ファイルを取得する](/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="b18d8-p140">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="b18d8-533">アプリが閲覧モードで `getCallbackTokenAsync` メソッドを呼び出すには、アプリのマニフェスト内に **ReadItem** アクセス許可が指定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="b18d8-533">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="b18d8-p141">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出してアイテムの識別子を `getCallbackTokenAsync` メソッドに渡す必要があります。アプリには、`saveAsync` メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="b18d8-p141">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b18d8-536">パラメーター</span><span class="sxs-lookup"><span data-stu-id="b18d8-536">Parameters</span></span>

|<span data-ttu-id="b18d8-537">名前</span><span class="sxs-lookup"><span data-stu-id="b18d8-537">Name</span></span>| <span data-ttu-id="b18d8-538">型</span><span class="sxs-lookup"><span data-stu-id="b18d8-538">Type</span></span>| <span data-ttu-id="b18d8-539">属性</span><span class="sxs-lookup"><span data-stu-id="b18d8-539">Attributes</span></span>| <span data-ttu-id="b18d8-540">説明</span><span class="sxs-lookup"><span data-stu-id="b18d8-540">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="b18d8-541">関数</span><span class="sxs-lookup"><span data-stu-id="b18d8-541">function</span></span>||<span data-ttu-id="b18d8-542">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="b18d8-542">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="b18d8-543">トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="b18d8-543">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="b18d8-544">エラーが発生した場合は`asyncResult.error` 、 `asyncResult.diagnostics`プロパティとプロパティによって追加情報が提供されることがあります。</span><span class="sxs-lookup"><span data-stu-id="b18d8-544">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="b18d8-545">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="b18d8-545">Object</span></span>| <span data-ttu-id="b18d8-546">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="b18d8-546">&lt;optional&gt;</span></span>|<span data-ttu-id="b18d8-547">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="b18d8-547">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="b18d8-548">エラー</span><span class="sxs-lookup"><span data-stu-id="b18d8-548">Errors</span></span>

|<span data-ttu-id="b18d8-549">エラー コード</span><span class="sxs-lookup"><span data-stu-id="b18d8-549">Error code</span></span>|<span data-ttu-id="b18d8-550">説明</span><span class="sxs-lookup"><span data-stu-id="b18d8-550">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="b18d8-551">要求が失敗しました。</span><span class="sxs-lookup"><span data-stu-id="b18d8-551">The request has failed.</span></span> <span data-ttu-id="b18d8-552">HTTP エラーコードについては、diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="b18d8-552">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="b18d8-553">Exchange サーバーがエラーを返しました。</span><span class="sxs-lookup"><span data-stu-id="b18d8-553">The Exchange server returned an error.</span></span> <span data-ttu-id="b18d8-554">詳細については、「diagnostics オブジェクト」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="b18d8-554">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="b18d8-555">ユーザーがネットワークに接続されていません。</span><span class="sxs-lookup"><span data-stu-id="b18d8-555">The user is no longer connected to the network.</span></span> <span data-ttu-id="b18d8-556">ネットワーク接続を確認し、もう一度実行してください。</span><span class="sxs-lookup"><span data-stu-id="b18d8-556">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b18d8-557">要件</span><span class="sxs-lookup"><span data-stu-id="b18d8-557">Requirements</span></span>

|<span data-ttu-id="b18d8-558">要件</span><span class="sxs-lookup"><span data-stu-id="b18d8-558">Requirement</span></span>| <span data-ttu-id="b18d8-559">値</span><span class="sxs-lookup"><span data-stu-id="b18d8-559">Value</span></span>|
|---|---|
|[<span data-ttu-id="b18d8-560">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b18d8-560">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b18d8-561">1.0</span><span class="sxs-lookup"><span data-stu-id="b18d8-561">1.0</span></span>|
|[<span data-ttu-id="b18d8-562">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b18d8-562">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b18d8-563">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b18d8-563">ReadItem</span></span>|
|[<span data-ttu-id="b18d8-564">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b18d8-564">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b18d8-565">新規作成と閲覧</span><span class="sxs-lookup"><span data-stu-id="b18d8-565">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="b18d8-566">例</span><span class="sxs-lookup"><span data-stu-id="b18d8-566">Example</span></span>

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

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="b18d8-567">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="b18d8-567">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="b18d8-568">ユーザーと Office アドインを識別するトークンを取得します。</span><span class="sxs-lookup"><span data-stu-id="b18d8-568">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="b18d8-569">`getUserIdentityTokenAsync` メソッドは、[アドインとユーザーをサード パーティのシステムで識別して認証](/outlook/add-ins/authentication)することのできるトークンを返します。</span><span class="sxs-lookup"><span data-stu-id="b18d8-569">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="b18d8-570">パラメーター</span><span class="sxs-lookup"><span data-stu-id="b18d8-570">Parameters</span></span>

|<span data-ttu-id="b18d8-571">名前</span><span class="sxs-lookup"><span data-stu-id="b18d8-571">Name</span></span>| <span data-ttu-id="b18d8-572">型</span><span class="sxs-lookup"><span data-stu-id="b18d8-572">Type</span></span>| <span data-ttu-id="b18d8-573">属性</span><span class="sxs-lookup"><span data-stu-id="b18d8-573">Attributes</span></span>| <span data-ttu-id="b18d8-574">説明</span><span class="sxs-lookup"><span data-stu-id="b18d8-574">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="b18d8-575">関数</span><span class="sxs-lookup"><span data-stu-id="b18d8-575">function</span></span>||<span data-ttu-id="b18d8-576">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="b18d8-576">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="b18d8-577">トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="b18d8-577">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="b18d8-578">エラーが発生した場合は`asyncResult.error` 、 `asyncResult.diagnostics`プロパティとプロパティによって追加情報が提供されることがあります。</span><span class="sxs-lookup"><span data-stu-id="b18d8-578">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="b18d8-579">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="b18d8-579">Object</span></span>| <span data-ttu-id="b18d8-580">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="b18d8-580">&lt;optional&gt;</span></span>|<span data-ttu-id="b18d8-581">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="b18d8-581">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="b18d8-582">エラー</span><span class="sxs-lookup"><span data-stu-id="b18d8-582">Errors</span></span>

|<span data-ttu-id="b18d8-583">エラー コード</span><span class="sxs-lookup"><span data-stu-id="b18d8-583">Error code</span></span>|<span data-ttu-id="b18d8-584">説明</span><span class="sxs-lookup"><span data-stu-id="b18d8-584">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="b18d8-585">要求が失敗しました。</span><span class="sxs-lookup"><span data-stu-id="b18d8-585">The request has failed.</span></span> <span data-ttu-id="b18d8-586">HTTP エラーコードについては、diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="b18d8-586">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="b18d8-587">Exchange サーバーがエラーを返しました。</span><span class="sxs-lookup"><span data-stu-id="b18d8-587">The Exchange server returned an error.</span></span> <span data-ttu-id="b18d8-588">詳細については、「diagnostics オブジェクト」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="b18d8-588">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="b18d8-589">ユーザーがネットワークに接続されていません。</span><span class="sxs-lookup"><span data-stu-id="b18d8-589">The user is no longer connected to the network.</span></span> <span data-ttu-id="b18d8-590">ネットワーク接続を確認し、もう一度実行してください。</span><span class="sxs-lookup"><span data-stu-id="b18d8-590">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b18d8-591">要件</span><span class="sxs-lookup"><span data-stu-id="b18d8-591">Requirements</span></span>

|<span data-ttu-id="b18d8-592">要件</span><span class="sxs-lookup"><span data-stu-id="b18d8-592">Requirement</span></span>| <span data-ttu-id="b18d8-593">値</span><span class="sxs-lookup"><span data-stu-id="b18d8-593">Value</span></span>|
|---|---|
|[<span data-ttu-id="b18d8-594">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b18d8-594">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b18d8-595">1.0</span><span class="sxs-lookup"><span data-stu-id="b18d8-595">1.0</span></span>|
|[<span data-ttu-id="b18d8-596">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b18d8-596">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b18d8-597">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b18d8-597">ReadItem</span></span>|
|[<span data-ttu-id="b18d8-598">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b18d8-598">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b18d8-599">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="b18d8-599">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b18d8-600">例</span><span class="sxs-lookup"><span data-stu-id="b18d8-600">Example</span></span>

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

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="b18d8-601">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="b18d8-601">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="b18d8-602">ユーザーのメールボックスをホストしている Exchange サーバー上の Exchange Web サービス (EWS) のサービスに対して非同期の要求を行います。</span><span class="sxs-lookup"><span data-stu-id="b18d8-602">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="b18d8-603">このメソッドは、次のシナリオではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="b18d8-603">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="b18d8-604">Outlook on iOS または Android</span><span class="sxs-lookup"><span data-stu-id="b18d8-604">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="b18d8-605">アドインが Gmail のメールボックスに読み込まれる場合</span><span class="sxs-lookup"><span data-stu-id="b18d8-605">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="b18d8-606">このような場合は、アドインでは [REST API を使用](/outlook/add-ins/use-rest-api)して、代わりにユーザーのメールボックスにアクセスする必要があります。</span><span class="sxs-lookup"><span data-stu-id="b18d8-606">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="b18d8-p148">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange. See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span><span class="sxs-lookup"><span data-stu-id="b18d8-p148">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange. See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="b18d8-609">`makeEwsRequestAsync` メソッドでは、フォルダー関連アイテムを要求できません。</span><span class="sxs-lookup"><span data-stu-id="b18d8-609">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="b18d8-610">XML 要求では UTF-8 エンコードを指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="b18d8-610">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="b18d8-p149">`makeEwsRequestAsync` メソッドを使用するには、アドインに **ReadWriteMailbox** アクセス許可が必要です。**ReadWriteMailbox** アクセス許可と、`makeEwsRequestAsync` メソッドで呼び出せる EWS 操作の使い方については、「[ユーザーのメールボックスへのメール アドイン アクセスのアクセス許可を指定する](/outlook/add-ins/understanding-outlook-add-in-permissions)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="b18d8-p149">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="b18d8-613">サーバー管理者は、クライアント アクセス サーバーの EWS ディレクトリで `OAuthAuthentication` を true に設定して、`makeEwsRequestAsync` メソッドで EWS 要求を行うことができるようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="b18d8-613">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="b18d8-614">バージョンの相違点</span><span class="sxs-lookup"><span data-stu-id="b18d8-614">Version differences</span></span>

<span data-ttu-id="b18d8-615">バージョン 15.0.4535.1004 より前のバージョンの Outlook で実行しているメール アプリで `makeEwsRequestAsync` メソッドを使う場合は、エンコード値を `ISO-8859-1` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="b18d8-615">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="b18d8-p150">Outlook on the web でメール アプリを実行している場合は、エンコード値を設定する必要はありません。mailbox.diagnostics.hostName プロパティを使って、メール アプリを Outlook で実行しているのか、Outlook on the web で実行しているのかを確認できます。mailbox.diagnostics.hostVersion プロパティを使って、どのバージョンの Outlook を使って実行しているかを確認できます。</span><span class="sxs-lookup"><span data-stu-id="b18d8-p150">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b18d8-619">パラメーター</span><span class="sxs-lookup"><span data-stu-id="b18d8-619">Parameters</span></span>

|<span data-ttu-id="b18d8-620">名前</span><span class="sxs-lookup"><span data-stu-id="b18d8-620">Name</span></span>| <span data-ttu-id="b18d8-621">型</span><span class="sxs-lookup"><span data-stu-id="b18d8-621">Type</span></span>| <span data-ttu-id="b18d8-622">属性</span><span class="sxs-lookup"><span data-stu-id="b18d8-622">Attributes</span></span>| <span data-ttu-id="b18d8-623">説明</span><span class="sxs-lookup"><span data-stu-id="b18d8-623">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="b18d8-624">String</span><span class="sxs-lookup"><span data-stu-id="b18d8-624">String</span></span>||<span data-ttu-id="b18d8-625">EWS 要求です。</span><span class="sxs-lookup"><span data-stu-id="b18d8-625">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="b18d8-626">function</span><span class="sxs-lookup"><span data-stu-id="b18d8-626">function</span></span>||<span data-ttu-id="b18d8-627">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="b18d8-627">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="b18d8-p151">The XML result of the EWS call is provided as a string in the `asyncResult.value` property. If the result exceeds 1 MB in size, an error message is returned instead.</span><span class="sxs-lookup"><span data-stu-id="b18d8-p151">The XML result of the EWS call is provided as a string in the `asyncResult.value` property. If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="b18d8-630">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="b18d8-630">Object</span></span>| <span data-ttu-id="b18d8-631">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="b18d8-631">&lt;optional&gt;</span></span>|<span data-ttu-id="b18d8-632">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="b18d8-632">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b18d8-633">要件</span><span class="sxs-lookup"><span data-stu-id="b18d8-633">Requirements</span></span>

|<span data-ttu-id="b18d8-634">要件</span><span class="sxs-lookup"><span data-stu-id="b18d8-634">Requirement</span></span>| <span data-ttu-id="b18d8-635">値</span><span class="sxs-lookup"><span data-stu-id="b18d8-635">Value</span></span>|
|---|---|
|[<span data-ttu-id="b18d8-636">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b18d8-636">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b18d8-637">1.0</span><span class="sxs-lookup"><span data-stu-id="b18d8-637">1.0</span></span>|
|[<span data-ttu-id="b18d8-638">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b18d8-638">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b18d8-639">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="b18d8-639">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="b18d8-640">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b18d8-640">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b18d8-641">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="b18d8-641">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b18d8-642">例</span><span class="sxs-lookup"><span data-stu-id="b18d8-642">Example</span></span>

<span data-ttu-id="b18d8-643">次の例は、`GetItem` 操作を使ってアイテムの件名を取得するため、`makeEwsRequestAsync` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="b18d8-643">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="b18d8-644">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="b18d8-644">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="b18d8-645">サポートされているイベントの種類のイベント ハンドラーを削除します。</span><span class="sxs-lookup"><span data-stu-id="b18d8-645">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="b18d8-646">現在、サポートされているイベントの種類は `Office.EventType.ItemChanged` だけです。</span><span class="sxs-lookup"><span data-stu-id="b18d8-646">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b18d8-647">パラメーター</span><span class="sxs-lookup"><span data-stu-id="b18d8-647">Parameters</span></span>

| <span data-ttu-id="b18d8-648">名前</span><span class="sxs-lookup"><span data-stu-id="b18d8-648">Name</span></span> | <span data-ttu-id="b18d8-649">型</span><span class="sxs-lookup"><span data-stu-id="b18d8-649">Type</span></span> | <span data-ttu-id="b18d8-650">属性</span><span class="sxs-lookup"><span data-stu-id="b18d8-650">Attributes</span></span> | <span data-ttu-id="b18d8-651">説明</span><span class="sxs-lookup"><span data-stu-id="b18d8-651">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="b18d8-652">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="b18d8-652">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="b18d8-653">ハンドラーを取り消すイベント。</span><span class="sxs-lookup"><span data-stu-id="b18d8-653">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="b18d8-654">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="b18d8-654">Object</span></span> | <span data-ttu-id="b18d8-655">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="b18d8-655">&lt;optional&gt;</span></span> | <span data-ttu-id="b18d8-656">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="b18d8-656">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="b18d8-657">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="b18d8-657">Object</span></span> | <span data-ttu-id="b18d8-658">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="b18d8-658">&lt;optional&gt;</span></span> | <span data-ttu-id="b18d8-659">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="b18d8-659">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="b18d8-660">function</span><span class="sxs-lookup"><span data-stu-id="b18d8-660">function</span></span>| <span data-ttu-id="b18d8-661">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b18d8-661">&lt;optional&gt;</span></span>|<span data-ttu-id="b18d8-662">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="b18d8-662">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b18d8-663">要件</span><span class="sxs-lookup"><span data-stu-id="b18d8-663">Requirements</span></span>

|<span data-ttu-id="b18d8-664">要件</span><span class="sxs-lookup"><span data-stu-id="b18d8-664">Requirement</span></span>| <span data-ttu-id="b18d8-665">値</span><span class="sxs-lookup"><span data-stu-id="b18d8-665">Value</span></span>|
|---|---|
|[<span data-ttu-id="b18d8-666">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b18d8-666">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b18d8-667">1.5</span><span class="sxs-lookup"><span data-stu-id="b18d8-667">1.5</span></span> |
|[<span data-ttu-id="b18d8-668">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b18d8-668">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b18d8-669">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b18d8-669">ReadItem</span></span> |
|[<span data-ttu-id="b18d8-670">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b18d8-670">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b18d8-671">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="b18d8-671">Compose or Read</span></span>|
