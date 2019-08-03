---
title: Office.context.mailbox - 要件セット 1.5
description: ''
ms.date: 04/24/2019
localization_priority: Priority
ms.openlocfilehash: 9ffb0d4d33af80a669fd81bc0130f14f673e9400
ms.sourcegitcommit: 3f5d7f4794e3d3c8bc3a79fa05c54157613b9376
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/02/2019
ms.locfileid: "36064754"
---
# <a name="mailbox"></a><span data-ttu-id="c91fe-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="c91fe-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="c91fe-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="c91fe-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="c91fe-104">Microsoft Outlook の Outlook アドイン オブジェクト モデルへのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="c91fe-104">Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c91fe-105">要件</span><span class="sxs-lookup"><span data-stu-id="c91fe-105">Requirements</span></span>

|<span data-ttu-id="c91fe-106">要件</span><span class="sxs-lookup"><span data-stu-id="c91fe-106">Requirement</span></span>| <span data-ttu-id="c91fe-107">値</span><span class="sxs-lookup"><span data-stu-id="c91fe-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="c91fe-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c91fe-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c91fe-109">1.0</span><span class="sxs-lookup"><span data-stu-id="c91fe-109">1.0</span></span>|
|[<span data-ttu-id="c91fe-110">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c91fe-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c91fe-111">制限あり</span><span class="sxs-lookup"><span data-stu-id="c91fe-111">Restricted</span></span>|
|[<span data-ttu-id="c91fe-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c91fe-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c91fe-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c91fe-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="c91fe-114">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="c91fe-114">Members and methods</span></span>

| <span data-ttu-id="c91fe-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="c91fe-115">Member</span></span> | <span data-ttu-id="c91fe-116">種類</span><span class="sxs-lookup"><span data-stu-id="c91fe-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="c91fe-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="c91fe-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="c91fe-118">メンバー</span><span class="sxs-lookup"><span data-stu-id="c91fe-118">Member</span></span> |
| [<span data-ttu-id="c91fe-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="c91fe-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="c91fe-120">メンバー</span><span class="sxs-lookup"><span data-stu-id="c91fe-120">Member</span></span> |
| [<span data-ttu-id="c91fe-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="c91fe-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="c91fe-122">メソッド</span><span class="sxs-lookup"><span data-stu-id="c91fe-122">Method</span></span> |
| [<span data-ttu-id="c91fe-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="c91fe-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="c91fe-124">メソッド</span><span class="sxs-lookup"><span data-stu-id="c91fe-124">Method</span></span> |
| [<span data-ttu-id="c91fe-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="c91fe-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="c91fe-126">メソッド</span><span class="sxs-lookup"><span data-stu-id="c91fe-126">Method</span></span> |
| [<span data-ttu-id="c91fe-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="c91fe-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="c91fe-128">メソッド</span><span class="sxs-lookup"><span data-stu-id="c91fe-128">Method</span></span> |
| [<span data-ttu-id="c91fe-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="c91fe-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="c91fe-130">メソッド</span><span class="sxs-lookup"><span data-stu-id="c91fe-130">Method</span></span> |
| [<span data-ttu-id="c91fe-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="c91fe-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="c91fe-132">メソッド</span><span class="sxs-lookup"><span data-stu-id="c91fe-132">Method</span></span> |
| [<span data-ttu-id="c91fe-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="c91fe-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="c91fe-134">メソッド</span><span class="sxs-lookup"><span data-stu-id="c91fe-134">Method</span></span> |
| [<span data-ttu-id="c91fe-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="c91fe-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="c91fe-136">メソッド</span><span class="sxs-lookup"><span data-stu-id="c91fe-136">Method</span></span> |
| [<span data-ttu-id="c91fe-137">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="c91fe-137">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="c91fe-138">メソッド</span><span class="sxs-lookup"><span data-stu-id="c91fe-138">Method</span></span> |
| [<span data-ttu-id="c91fe-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="c91fe-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="c91fe-140">メソッド</span><span class="sxs-lookup"><span data-stu-id="c91fe-140">Method</span></span> |
| [<span data-ttu-id="c91fe-141">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="c91fe-141">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="c91fe-142">メソッド</span><span class="sxs-lookup"><span data-stu-id="c91fe-142">Method</span></span> |
| [<span data-ttu-id="c91fe-143">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="c91fe-143">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="c91fe-144">メソッド</span><span class="sxs-lookup"><span data-stu-id="c91fe-144">Method</span></span> |
| [<span data-ttu-id="c91fe-145">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="c91fe-145">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="c91fe-146">メソッド</span><span class="sxs-lookup"><span data-stu-id="c91fe-146">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="c91fe-147">名前空間</span><span class="sxs-lookup"><span data-stu-id="c91fe-147">Namespaces</span></span>

<span data-ttu-id="c91fe-148">[diagnostics](Office.context.mailbox.diagnostics.md):Outlook アドインに診断情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="c91fe-148">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="c91fe-149">[item](Office.context.mailbox.item.md):Outlook アドインのメッセージや予定にアクセスするためのメソッドとプロパティを提供します。</span><span class="sxs-lookup"><span data-stu-id="c91fe-149">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="c91fe-150">[userProfile](Office.context.mailbox.userProfile.md):Outlook アドインのユーザーに関する情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="c91fe-150">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="c91fe-151">Members</span><span class="sxs-lookup"><span data-stu-id="c91fe-151">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="c91fe-152">ewsUrl: String</span><span class="sxs-lookup"><span data-stu-id="c91fe-152">ewsUrl :String</span></span>

<span data-ttu-id="c91fe-p101">このメール アカウントの Exchange Web サービス (EWS) エンドポイントの URL を取得します。読み取りモードのみです。</span><span class="sxs-lookup"><span data-stu-id="c91fe-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="c91fe-155">このメンバーは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c91fe-155">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c91fe-p102">
  `ewsUrl\` 値は、リモート サービスで、ユーザーのメールボックスに EWS 呼び出しを行うために使うことができます。たとえば、[選択したアイテムから添付ファイルを取得する](/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="c91fe-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="c91fe-158">アプリが閲覧モードで `ewsUrl` メンバーを呼び出すには、アプリのマニフェスト内に **ReadItem** アクセス許可が指定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="c91fe-158">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="c91fe-p103">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出してから、`ewsUrl` メンバーを使用する必要があります。アプリには、`saveAsync` メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="c91fe-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="c91fe-161">型</span><span class="sxs-lookup"><span data-stu-id="c91fe-161">Type</span></span>

*   <span data-ttu-id="c91fe-162">String</span><span class="sxs-lookup"><span data-stu-id="c91fe-162">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c91fe-163">要件</span><span class="sxs-lookup"><span data-stu-id="c91fe-163">Requirements</span></span>

|<span data-ttu-id="c91fe-164">要件</span><span class="sxs-lookup"><span data-stu-id="c91fe-164">Requirement</span></span>| <span data-ttu-id="c91fe-165">値</span><span class="sxs-lookup"><span data-stu-id="c91fe-165">Value</span></span>|
|---|---|
|[<span data-ttu-id="c91fe-166">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c91fe-166">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c91fe-167">1.0</span><span class="sxs-lookup"><span data-stu-id="c91fe-167">1.0</span></span>|
|[<span data-ttu-id="c91fe-168">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c91fe-168">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c91fe-169">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c91fe-169">ReadItem</span></span>|
|[<span data-ttu-id="c91fe-170">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c91fe-170">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c91fe-171">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c91fe-171">Compose or Read</span></span>|

#### <a name="resturl-string"></a><span data-ttu-id="c91fe-172">restUrl: String</span><span class="sxs-lookup"><span data-stu-id="c91fe-172">restUrl :String</span></span>

<span data-ttu-id="c91fe-173">この電子メール アカウントの REST エンドポイントの URL を取得します。</span><span class="sxs-lookup"><span data-stu-id="c91fe-173">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="c91fe-174">`restUrl` 値は、ユーザーのメールボックスに [REST API](/outlook/rest/) 呼び出しを行うために使用できます。</span><span class="sxs-lookup"><span data-stu-id="c91fe-174">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="c91fe-175">アプリが閲覧モードで `restUrl` メンバーを呼び出すには、アプリのマニフェスト内に **ReadItem** アクセス許可が指定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="c91fe-175">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="c91fe-p104">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出してから、`restUrl` メンバーを使用する必要があります。アプリには、`saveAsync` メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="c91fe-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

> [!NOTE]
> <span data-ttu-id="c91fe-178">構成されたカスタム REST URL を使用する Exchange 2016 以降のオンプレミスのインストールに接続されている Outlook クライアントは、`restUrl` に無効な値を返します。</span><span class="sxs-lookup"><span data-stu-id="c91fe-178">Outlook clients connected to on-premises installations of Exchange 2016 or later with a custom REST URL configured will return an invalid value for `restUrl`.</span></span>

##### <a name="type"></a><span data-ttu-id="c91fe-179">型</span><span class="sxs-lookup"><span data-stu-id="c91fe-179">Type</span></span>

*   <span data-ttu-id="c91fe-180">String</span><span class="sxs-lookup"><span data-stu-id="c91fe-180">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c91fe-181">要件</span><span class="sxs-lookup"><span data-stu-id="c91fe-181">Requirements</span></span>

|<span data-ttu-id="c91fe-182">要件</span><span class="sxs-lookup"><span data-stu-id="c91fe-182">Requirement</span></span>| <span data-ttu-id="c91fe-183">値</span><span class="sxs-lookup"><span data-stu-id="c91fe-183">Value</span></span>|
|---|---|
|[<span data-ttu-id="c91fe-184">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c91fe-184">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c91fe-185">1.5</span><span class="sxs-lookup"><span data-stu-id="c91fe-185">1.5</span></span> |
|[<span data-ttu-id="c91fe-186">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c91fe-186">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c91fe-187">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c91fe-187">ReadItem</span></span>|
|[<span data-ttu-id="c91fe-188">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c91fe-188">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c91fe-189">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c91fe-189">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="c91fe-190">メソッド</span><span class="sxs-lookup"><span data-stu-id="c91fe-190">Methods</span></span>

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="c91fe-191">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c91fe-191">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="c91fe-192">サポートされているイベントのイベント ハンドラーを追加します。</span><span class="sxs-lookup"><span data-stu-id="c91fe-192">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="c91fe-193">現在サポートされている唯一のイベントの種類は `Office.EventType.ItemChanged` です。これはユーザーが新しいアイテムを選択したときに呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c91fe-193">Currently the only supported event type is `Office.EventType.ItemChanged`, which is invoked when the user selects a new item.</span></span> <span data-ttu-id="c91fe-194">このイベントは、ピン留め可能な作業ウィンドウを実装するアドインで使用され、現在選択されているアイテムに基づいて作業ウィンドウ UI をアドインで更新できるようにします。</span><span class="sxs-lookup"><span data-stu-id="c91fe-194">This event is used by add-ins that implement a pinnable task pane, and allows the add-in to refresh the task pane UI based on the currently selected item.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c91fe-195">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c91fe-195">Parameters</span></span>

| <span data-ttu-id="c91fe-196">名前</span><span class="sxs-lookup"><span data-stu-id="c91fe-196">Name</span></span> | <span data-ttu-id="c91fe-197">型</span><span class="sxs-lookup"><span data-stu-id="c91fe-197">Type</span></span> | <span data-ttu-id="c91fe-198">属性</span><span class="sxs-lookup"><span data-stu-id="c91fe-198">Attributes</span></span> | <span data-ttu-id="c91fe-199">説明</span><span class="sxs-lookup"><span data-stu-id="c91fe-199">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="c91fe-200">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="c91fe-200">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="c91fe-201">ハンドラーを呼び出す必要のあるイベント。</span><span class="sxs-lookup"><span data-stu-id="c91fe-201">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="c91fe-202">Function</span><span class="sxs-lookup"><span data-stu-id="c91fe-202">Function</span></span> || <span data-ttu-id="c91fe-p106">イベントを処理する関数。関数は、オブジェクト リテラルである単一パラメーターを受け入れる必要があります。パラメーターの `type` プロパティは、`addHandlerAsync` に渡される `eventType` パラメーターと一致します。</span><span class="sxs-lookup"><span data-stu-id="c91fe-p106">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="c91fe-206">Object</span><span class="sxs-lookup"><span data-stu-id="c91fe-206">Object</span></span> | <span data-ttu-id="c91fe-207">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c91fe-207">&lt;optional&gt;</span></span> | <span data-ttu-id="c91fe-208">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="c91fe-208">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="c91fe-209">Object</span><span class="sxs-lookup"><span data-stu-id="c91fe-209">Object</span></span> | <span data-ttu-id="c91fe-210">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c91fe-210">&lt;optional&gt;</span></span> | <span data-ttu-id="c91fe-211">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="c91fe-211">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="c91fe-212">function</span><span class="sxs-lookup"><span data-stu-id="c91fe-212">function</span></span>| <span data-ttu-id="c91fe-213">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c91fe-213">&lt;optional&gt;</span></span>|<span data-ttu-id="c91fe-214">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c91fe-214">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c91fe-215">要件</span><span class="sxs-lookup"><span data-stu-id="c91fe-215">Requirements</span></span>

|<span data-ttu-id="c91fe-216">要件</span><span class="sxs-lookup"><span data-stu-id="c91fe-216">Requirement</span></span>| <span data-ttu-id="c91fe-217">値</span><span class="sxs-lookup"><span data-stu-id="c91fe-217">Value</span></span>|
|---|---|
|[<span data-ttu-id="c91fe-218">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c91fe-218">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c91fe-219">1.5</span><span class="sxs-lookup"><span data-stu-id="c91fe-219">1.5</span></span> |
|[<span data-ttu-id="c91fe-220">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c91fe-220">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c91fe-221">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c91fe-221">ReadItem</span></span> |
|[<span data-ttu-id="c91fe-222">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c91fe-222">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c91fe-223">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c91fe-223">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c91fe-224">例</span><span class="sxs-lookup"><span data-stu-id="c91fe-224">Example</span></span>

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

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="c91fe-225">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="c91fe-225">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="c91fe-226">REST 形式のアイテム ID を EWS 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="c91fe-226">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="c91fe-227">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c91fe-227">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c91fe-p107">REST API ([Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) や [Microsoft Graph](https://graph.microsoft.io/) など) で取得されたアイテム ID は、Exchange Web サービス (EWS) に使用される形式とは異なる形式を使用します。`convertToEwsId` メソッドは、REST 形式の ID を EWS 用の適切な形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="c91fe-p107">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c91fe-230">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c91fe-230">Parameters</span></span>

|<span data-ttu-id="c91fe-231">名前</span><span class="sxs-lookup"><span data-stu-id="c91fe-231">Name</span></span>| <span data-ttu-id="c91fe-232">種類</span><span class="sxs-lookup"><span data-stu-id="c91fe-232">Type</span></span>| <span data-ttu-id="c91fe-233">説明</span><span class="sxs-lookup"><span data-stu-id="c91fe-233">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="c91fe-234">String</span><span class="sxs-lookup"><span data-stu-id="c91fe-234">String</span></span>|<span data-ttu-id="c91fe-235">Outlook REST API 形式のアイテム ID</span><span class="sxs-lookup"><span data-stu-id="c91fe-235">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="c91fe-236">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="c91fe-236">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.5)|<span data-ttu-id="c91fe-237">アイテム ID の取得に使用された Outlook REST API のバージョンを示す値。</span><span class="sxs-lookup"><span data-stu-id="c91fe-237">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c91fe-238">要件</span><span class="sxs-lookup"><span data-stu-id="c91fe-238">Requirements</span></span>

|<span data-ttu-id="c91fe-239">要件</span><span class="sxs-lookup"><span data-stu-id="c91fe-239">Requirement</span></span>| <span data-ttu-id="c91fe-240">値</span><span class="sxs-lookup"><span data-stu-id="c91fe-240">Value</span></span>|
|---|---|
|[<span data-ttu-id="c91fe-241">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c91fe-241">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c91fe-242">1.3</span><span class="sxs-lookup"><span data-stu-id="c91fe-242">1.3</span></span>|
|[<span data-ttu-id="c91fe-243">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c91fe-243">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c91fe-244">制限あり</span><span class="sxs-lookup"><span data-stu-id="c91fe-244">Restricted</span></span>|
|[<span data-ttu-id="c91fe-245">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c91fe-245">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c91fe-246">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c91fe-246">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c91fe-247">戻り値:</span><span class="sxs-lookup"><span data-stu-id="c91fe-247">Returns:</span></span>

<span data-ttu-id="c91fe-248">型:String</span><span class="sxs-lookup"><span data-stu-id="c91fe-248">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="c91fe-249">例</span><span class="sxs-lookup"><span data-stu-id="c91fe-249">Example</span></span>

```javascript
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-15"></a><span data-ttu-id="c91fe-250">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.5)}</span><span class="sxs-lookup"><span data-stu-id="c91fe-250">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.5)}</span></span>

<span data-ttu-id="c91fe-251">クライアントのローカル時間で時間情報が含まれている辞書を取得します。</span><span class="sxs-lookup"><span data-stu-id="c91fe-251">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="c91fe-p108">Outlook on the web または Outlook デスクトップのメールアプリでは、日付と時刻に異なるタイムゾーンを使用できます。Outlook デスクトップは、クライアント コンピューターのタイム ゾーンを使用します。Outlook on the web は、Exchange 管理センター (EAC) で設定されたタイム ゾーンを使用します。日付と時刻の値は、ユーザー インターフェイスに表示される値が、常にユーザーが期待するタイム ゾーンと一致するように処理する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c91fe-p108">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="c91fe-p109">Outlook デスクトップ クライアントでメール アプリを実行している場合、`convertToLocalClientTime` メソッドは、クライアント コンピューターのタイム ゾーンに設定された値のディクショナリ オブジェクトを返します。Outlook on the web でメール アプリを実行している場合、`convertToLocalClientTime` メソッドは、EAC で指定したタイム ゾーンに設定された値のディクショナリ オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c91fe-p109">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c91fe-257">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c91fe-257">Parameters</span></span>

|<span data-ttu-id="c91fe-258">名前</span><span class="sxs-lookup"><span data-stu-id="c91fe-258">Name</span></span>| <span data-ttu-id="c91fe-259">種類</span><span class="sxs-lookup"><span data-stu-id="c91fe-259">Type</span></span>| <span data-ttu-id="c91fe-260">説明</span><span class="sxs-lookup"><span data-stu-id="c91fe-260">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="c91fe-261">Date</span><span class="sxs-lookup"><span data-stu-id="c91fe-261">Date</span></span>|<span data-ttu-id="c91fe-262">日付オブジェクト</span><span class="sxs-lookup"><span data-stu-id="c91fe-262">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c91fe-263">要件</span><span class="sxs-lookup"><span data-stu-id="c91fe-263">Requirements</span></span>

|<span data-ttu-id="c91fe-264">要件</span><span class="sxs-lookup"><span data-stu-id="c91fe-264">Requirement</span></span>| <span data-ttu-id="c91fe-265">値</span><span class="sxs-lookup"><span data-stu-id="c91fe-265">Value</span></span>|
|---|---|
|[<span data-ttu-id="c91fe-266">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c91fe-266">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c91fe-267">1.0</span><span class="sxs-lookup"><span data-stu-id="c91fe-267">1.0</span></span>|
|[<span data-ttu-id="c91fe-268">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c91fe-268">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c91fe-269">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c91fe-269">ReadItem</span></span>|
|[<span data-ttu-id="c91fe-270">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c91fe-270">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c91fe-271">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c91fe-271">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c91fe-272">戻り値:</span><span class="sxs-lookup"><span data-stu-id="c91fe-272">Returns:</span></span>

<span data-ttu-id="c91fe-273">型:[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="c91fe-273">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.5)</span></span>

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="c91fe-274">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="c91fe-274">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="c91fe-275">EWS 形式のアイテム ID を REST 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="c91fe-275">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="c91fe-276">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c91fe-276">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c91fe-p110">EWS または `itemId` プロパティで取得されるアイテム ID は、REST API ([Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) や [Microsoft Graph](https://graph.microsoft.io/) など) に使用される形式とは異なる形式を使用します。`convertToRestId` メソッドは、EWS 形式の ID を REST 用の適切な形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="c91fe-p110">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c91fe-279">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c91fe-279">Parameters</span></span>

|<span data-ttu-id="c91fe-280">名前</span><span class="sxs-lookup"><span data-stu-id="c91fe-280">Name</span></span>| <span data-ttu-id="c91fe-281">種類</span><span class="sxs-lookup"><span data-stu-id="c91fe-281">Type</span></span>| <span data-ttu-id="c91fe-282">説明</span><span class="sxs-lookup"><span data-stu-id="c91fe-282">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="c91fe-283">String</span><span class="sxs-lookup"><span data-stu-id="c91fe-283">String</span></span>|<span data-ttu-id="c91fe-284">Exchange Web サービス (EWS) 形式のアイテム ID</span><span class="sxs-lookup"><span data-stu-id="c91fe-284">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="c91fe-285">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="c91fe-285">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.5)|<span data-ttu-id="c91fe-286">変換後の ID を使用する Outlook REST API のバージョンを示す値。</span><span class="sxs-lookup"><span data-stu-id="c91fe-286">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c91fe-287">要件</span><span class="sxs-lookup"><span data-stu-id="c91fe-287">Requirements</span></span>

|<span data-ttu-id="c91fe-288">要件</span><span class="sxs-lookup"><span data-stu-id="c91fe-288">Requirement</span></span>| <span data-ttu-id="c91fe-289">値</span><span class="sxs-lookup"><span data-stu-id="c91fe-289">Value</span></span>|
|---|---|
|[<span data-ttu-id="c91fe-290">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c91fe-290">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c91fe-291">1.3</span><span class="sxs-lookup"><span data-stu-id="c91fe-291">1.3</span></span>|
|[<span data-ttu-id="c91fe-292">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c91fe-292">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c91fe-293">制限あり</span><span class="sxs-lookup"><span data-stu-id="c91fe-293">Restricted</span></span>|
|[<span data-ttu-id="c91fe-294">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c91fe-294">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c91fe-295">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c91fe-295">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c91fe-296">戻り値:</span><span class="sxs-lookup"><span data-stu-id="c91fe-296">Returns:</span></span>

<span data-ttu-id="c91fe-297">型:String</span><span class="sxs-lookup"><span data-stu-id="c91fe-297">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="c91fe-298">例</span><span class="sxs-lookup"><span data-stu-id="c91fe-298">Example</span></span>

```javascript
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="c91fe-299">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="c91fe-299">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="c91fe-300">時間情報が含まれているディクショナリから日付オブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="c91fe-300">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="c91fe-301">`convertToUtcClientTime` メソッドは、ローカルの日付と時刻を含むディクショナリを、ローカルの日付と時刻の正しい値を持つ日付オブジェクトに変換します。</span><span class="sxs-lookup"><span data-stu-id="c91fe-301">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c91fe-302">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c91fe-302">Parameters</span></span>

|<span data-ttu-id="c91fe-303">名前</span><span class="sxs-lookup"><span data-stu-id="c91fe-303">Name</span></span>| <span data-ttu-id="c91fe-304">種類</span><span class="sxs-lookup"><span data-stu-id="c91fe-304">Type</span></span>| <span data-ttu-id="c91fe-305">説明</span><span class="sxs-lookup"><span data-stu-id="c91fe-305">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="c91fe-306">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="c91fe-306">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.5)|<span data-ttu-id="c91fe-307">変換するローカル時刻の値。</span><span class="sxs-lookup"><span data-stu-id="c91fe-307">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c91fe-308">要件</span><span class="sxs-lookup"><span data-stu-id="c91fe-308">Requirements</span></span>

|<span data-ttu-id="c91fe-309">要件</span><span class="sxs-lookup"><span data-stu-id="c91fe-309">Requirement</span></span>| <span data-ttu-id="c91fe-310">値</span><span class="sxs-lookup"><span data-stu-id="c91fe-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="c91fe-311">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c91fe-311">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c91fe-312">1.0</span><span class="sxs-lookup"><span data-stu-id="c91fe-312">1.0</span></span>|
|[<span data-ttu-id="c91fe-313">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c91fe-313">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c91fe-314">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c91fe-314">ReadItem</span></span>|
|[<span data-ttu-id="c91fe-315">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c91fe-315">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c91fe-316">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c91fe-316">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c91fe-317">戻り値:</span><span class="sxs-lookup"><span data-stu-id="c91fe-317">Returns:</span></span>

<span data-ttu-id="c91fe-318">時間が UTC で表現された日付オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="c91fe-318">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="c91fe-319">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="c91fe-319">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="c91fe-320">Date</span><span class="sxs-lookup"><span data-stu-id="c91fe-320">Date</span></span></dd>

</dl>

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="c91fe-321">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="c91fe-321">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="c91fe-322">既存の予定を表示します。</span><span class="sxs-lookup"><span data-stu-id="c91fe-322">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="c91fe-323">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c91fe-323">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c91fe-324">`displayAppointmentForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存の予定を開きます。</span><span class="sxs-lookup"><span data-stu-id="c91fe-324">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="c91fe-p111">Outlook on Mac では、このメソッドを使用して定期的な系列に含まれない単発の予定や定期的な系列のマスター予定を表示できます。ただし、系列のインスタンスは表示できません。これは、Outlook on Mac では定期的な系列のインスタンスのプロパティ (アイテム ID を含む) にアクセスできないためです。</span><span class="sxs-lookup"><span data-stu-id="c91fe-p111">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="c91fe-327">Outlook on the web では、このメソッドはフォームの本文が 32KB 以下の文字数の場合にのみ指定のフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="c91fe-327">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="c91fe-328">指定のアイテム識別子が既存の予定を表していない場合は、クライアント コンピューターまたはデバイスで空のウィンドウが開き、エラー メッセージは返されません。</span><span class="sxs-lookup"><span data-stu-id="c91fe-328">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c91fe-329">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c91fe-329">Parameters</span></span>

|<span data-ttu-id="c91fe-330">名前</span><span class="sxs-lookup"><span data-stu-id="c91fe-330">Name</span></span>| <span data-ttu-id="c91fe-331">種類</span><span class="sxs-lookup"><span data-stu-id="c91fe-331">Type</span></span>| <span data-ttu-id="c91fe-332">説明</span><span class="sxs-lookup"><span data-stu-id="c91fe-332">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="c91fe-333">String</span><span class="sxs-lookup"><span data-stu-id="c91fe-333">String</span></span>|<span data-ttu-id="c91fe-334">既存の予定の Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="c91fe-334">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c91fe-335">要件</span><span class="sxs-lookup"><span data-stu-id="c91fe-335">Requirements</span></span>

|<span data-ttu-id="c91fe-336">要件</span><span class="sxs-lookup"><span data-stu-id="c91fe-336">Requirement</span></span>| <span data-ttu-id="c91fe-337">値</span><span class="sxs-lookup"><span data-stu-id="c91fe-337">Value</span></span>|
|---|---|
|[<span data-ttu-id="c91fe-338">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c91fe-338">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c91fe-339">1.0</span><span class="sxs-lookup"><span data-stu-id="c91fe-339">1.0</span></span>|
|[<span data-ttu-id="c91fe-340">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c91fe-340">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c91fe-341">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c91fe-341">ReadItem</span></span>|
|[<span data-ttu-id="c91fe-342">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c91fe-342">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c91fe-343">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c91fe-343">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c91fe-344">例</span><span class="sxs-lookup"><span data-stu-id="c91fe-344">Example</span></span>

```javascript
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

#### <a name="displaymessageformitemid"></a><span data-ttu-id="c91fe-345">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="c91fe-345">displayMessageForm(itemId)</span></span>

<span data-ttu-id="c91fe-346">既存のメッセージを表示します。</span><span class="sxs-lookup"><span data-stu-id="c91fe-346">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="c91fe-347">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c91fe-347">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c91fe-348">`displayMessageForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存のメッセージを開きます。</span><span class="sxs-lookup"><span data-stu-id="c91fe-348">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="c91fe-349">Outlook on the web では、このメソッドはフォームの本文が 32 KB 以下の文字数の場合にのみ指定のフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="c91fe-349">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="c91fe-350">指定のアイテム識別子が既存のメッセージを表していない場合は、クライアント コンピューターにはメッセージは表示されず、エラー メッセージも返されません。</span><span class="sxs-lookup"><span data-stu-id="c91fe-350">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="c91fe-p112">予定を表す `itemId` を含む `displayMessageForm` を使用しないでください。既存の予定を表示するには、`displayAppointmentForm` メソッドを使用します。新しい予定を作成するフォームを表示するには、`displayNewAppointmentForm` メソッドを使用します。</span><span class="sxs-lookup"><span data-stu-id="c91fe-p112">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c91fe-353">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c91fe-353">Parameters</span></span>

|<span data-ttu-id="c91fe-354">名前</span><span class="sxs-lookup"><span data-stu-id="c91fe-354">Name</span></span>| <span data-ttu-id="c91fe-355">種類</span><span class="sxs-lookup"><span data-stu-id="c91fe-355">Type</span></span>| <span data-ttu-id="c91fe-356">説明</span><span class="sxs-lookup"><span data-stu-id="c91fe-356">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="c91fe-357">String</span><span class="sxs-lookup"><span data-stu-id="c91fe-357">String</span></span>|<span data-ttu-id="c91fe-358">既存のメッセージの Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="c91fe-358">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c91fe-359">要件</span><span class="sxs-lookup"><span data-stu-id="c91fe-359">Requirements</span></span>

|<span data-ttu-id="c91fe-360">要件</span><span class="sxs-lookup"><span data-stu-id="c91fe-360">Requirement</span></span>| <span data-ttu-id="c91fe-361">値</span><span class="sxs-lookup"><span data-stu-id="c91fe-361">Value</span></span>|
|---|---|
|[<span data-ttu-id="c91fe-362">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c91fe-362">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c91fe-363">1.0</span><span class="sxs-lookup"><span data-stu-id="c91fe-363">1.0</span></span>|
|[<span data-ttu-id="c91fe-364">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c91fe-364">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c91fe-365">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c91fe-365">ReadItem</span></span>|
|[<span data-ttu-id="c91fe-366">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c91fe-366">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c91fe-367">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c91fe-367">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c91fe-368">例</span><span class="sxs-lookup"><span data-stu-id="c91fe-368">Example</span></span>

```javascript
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="c91fe-369">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="c91fe-369">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="c91fe-370">新しい予定を作成するためのフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="c91fe-370">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="c91fe-371">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c91fe-371">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c91fe-p113">`displayNewAppointmentForm` メソッドを使用すると、ユーザーが新しい予定または会議を作成できるフォームが開きます。パラメーターを指定すると、予定のフォーム フィールドにパラメーターの内容が自動的に設定されます。</span><span class="sxs-lookup"><span data-stu-id="c91fe-p113">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="c91fe-p114">Outlook on the web およびモバイル デバイスでは、このメソッドは常に出席者フィールドが含まれるフォームを表示します。入力引数として出席者を指定しないと、このメソッドは **[保存]** ボタンのあるフォームを表示します。出席者を指定した場合には、フォームにその出席者と **[送信]** ボタンが表示されます。</span><span class="sxs-lookup"><span data-stu-id="c91fe-p114">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="c91fe-p115">Outlook リッチ クライアントと Outlook RT で、`requiredAttendees`、`optionalAttendees`、または `resources` パラメーターに出席者またはリソースを指定し、このメソッドを実行すると、**[送信]** ボタンがある会議フォームが表示されます。受信者を指定せずにこのメソッドを実行すると、**[保存して閉じる]** ボタンがある予定フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="c91fe-p115">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="c91fe-379">パラメーターのいずれかが指定のサイズ制限を超える場合、または不明なパラメーター名が指定されている場合は、例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="c91fe-379">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c91fe-380">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c91fe-380">Parameters</span></span>

|<span data-ttu-id="c91fe-381">名前</span><span class="sxs-lookup"><span data-stu-id="c91fe-381">Name</span></span>| <span data-ttu-id="c91fe-382">種類</span><span class="sxs-lookup"><span data-stu-id="c91fe-382">Type</span></span>| <span data-ttu-id="c91fe-383">説明</span><span class="sxs-lookup"><span data-stu-id="c91fe-383">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="c91fe-384">Object</span><span class="sxs-lookup"><span data-stu-id="c91fe-384">Object</span></span> | <span data-ttu-id="c91fe-385">新しい予定を記述するパラメーターのディクショナリ。</span><span class="sxs-lookup"><span data-stu-id="c91fe-385">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="c91fe-386">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)&gt;</span><span class="sxs-lookup"><span data-stu-id="c91fe-386">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)&gt;</span></span> | <span data-ttu-id="c91fe-p116">予定に必要な各出席者について、メール アドレスを含む文字列の配列、または `EmailAddressDetails` オブジェクトを含む配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="c91fe-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="c91fe-389">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)&gt;</span><span class="sxs-lookup"><span data-stu-id="c91fe-389">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)&gt;</span></span> | <span data-ttu-id="c91fe-p117">予定の各任意出席者について、メール アドレスを含む文字列の配列、または `EmailAddressDetails` オブジェクトを含む配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="c91fe-p117">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="c91fe-392">日付</span><span class="sxs-lookup"><span data-stu-id="c91fe-392">Date</span></span> | <span data-ttu-id="c91fe-393">予定の開始日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="c91fe-393">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="c91fe-394">Date</span><span class="sxs-lookup"><span data-stu-id="c91fe-394">Date</span></span> | <span data-ttu-id="c91fe-395">予定の終了日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="c91fe-395">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="c91fe-396">String</span><span class="sxs-lookup"><span data-stu-id="c91fe-396">String</span></span> | <span data-ttu-id="c91fe-p118">予定の場所を含む文字列。文字列は最大 255 文字に制限されます。</span><span class="sxs-lookup"><span data-stu-id="c91fe-p118">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="c91fe-399">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="c91fe-399">Array.&lt;String&gt;</span></span> | <span data-ttu-id="c91fe-p119">予定に必要なリソースを含む文字列の配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="c91fe-p119">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="c91fe-402">String</span><span class="sxs-lookup"><span data-stu-id="c91fe-402">String</span></span> | <span data-ttu-id="c91fe-p120">予定の件名を含む文字列です。文字列は最大 255 文字に制限されます。</span><span class="sxs-lookup"><span data-stu-id="c91fe-p120">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="c91fe-405">String</span><span class="sxs-lookup"><span data-stu-id="c91fe-405">String</span></span> | <span data-ttu-id="c91fe-p121">予定の本文。本文の内容は、最大サイズが 32 KB に制限されます。</span><span class="sxs-lookup"><span data-stu-id="c91fe-p121">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="c91fe-408">要件</span><span class="sxs-lookup"><span data-stu-id="c91fe-408">Requirements</span></span>

|<span data-ttu-id="c91fe-409">要件</span><span class="sxs-lookup"><span data-stu-id="c91fe-409">Requirement</span></span>| <span data-ttu-id="c91fe-410">値</span><span class="sxs-lookup"><span data-stu-id="c91fe-410">Value</span></span>|
|---|---|
|[<span data-ttu-id="c91fe-411">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c91fe-411">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c91fe-412">1.0</span><span class="sxs-lookup"><span data-stu-id="c91fe-412">1.0</span></span>|
|[<span data-ttu-id="c91fe-413">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c91fe-413">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c91fe-414">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c91fe-414">ReadItem</span></span>|
|[<span data-ttu-id="c91fe-415">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c91fe-415">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c91fe-416">読み取り</span><span class="sxs-lookup"><span data-stu-id="c91fe-416">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c91fe-417">例</span><span class="sxs-lookup"><span data-stu-id="c91fe-417">Example</span></span>

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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="c91fe-418">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="c91fe-418">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="c91fe-419">REST API または Exchange Web サービスを呼び出すために使用するトークンを含む文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="c91fe-419">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="c91fe-p122">`getCallbackTokenAsync` メソッドは、ユーザーのメールボックスをホストする Exchange Server から不透明なトークンを取得する非同期の呼び出しを行います。コールバック トークンの有効期間は 5 分です。</span><span class="sxs-lookup"><span data-stu-id="c91fe-p122">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="c91fe-422">可能であれば、アドインでは Exchange Web サービスの代わりに REST API を使用することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="c91fe-422">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span> 

<span data-ttu-id="c91fe-423">**REST トークン**</span><span class="sxs-lookup"><span data-stu-id="c91fe-423">**REST Tokens**</span></span>

<span data-ttu-id="c91fe-p123">REST トークンが要求された場合 (`options.isRest = true`)、結果トークンは Exchange Web サービスの呼び出しを認証するためには機能しません。アドインがマニフェストで [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) アクセス許可を指定していない限り、トークンの範囲は現在のアイテムとその添付ファイルへの読み取り専用アクセスに制限されます。`ReadWriteMailbox` アクセス許可が指定されている場合は、結果トークンは、メールを送信する機能など、メール、カレンダー、連絡先への読み取り/書き込みアクセスを付与します。</span><span class="sxs-lookup"><span data-stu-id="c91fe-p123">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="c91fe-427">アドインでは、`restUrl` プロパティを使用して、REST API 呼び出しを行うときに使用する正しい URL を決定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c91fe-427">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="c91fe-428">**EWS トークン**</span><span class="sxs-lookup"><span data-stu-id="c91fe-428">**EWS Tokens**</span></span>

<span data-ttu-id="c91fe-p124">EWS トークンが要求された場合 (`options.isRest = false`)、結果トークンは REST API 呼び出しを認証するためには機能しません。トークンの範囲は、現在のアイテムへのアクセスに制限されます。</span><span class="sxs-lookup"><span data-stu-id="c91fe-p124">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="c91fe-431">アドインでは、`ewsUrl` プロパティを使用して、EWS 呼び出しを行うときに使用する正しい URL を決定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c91fe-431">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c91fe-432">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c91fe-432">Parameters</span></span>

|<span data-ttu-id="c91fe-433">名前</span><span class="sxs-lookup"><span data-stu-id="c91fe-433">Name</span></span>| <span data-ttu-id="c91fe-434">型</span><span class="sxs-lookup"><span data-stu-id="c91fe-434">Type</span></span>| <span data-ttu-id="c91fe-435">属性</span><span class="sxs-lookup"><span data-stu-id="c91fe-435">Attributes</span></span>| <span data-ttu-id="c91fe-436">説明</span><span class="sxs-lookup"><span data-stu-id="c91fe-436">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="c91fe-437">Object</span><span class="sxs-lookup"><span data-stu-id="c91fe-437">Object</span></span> | <span data-ttu-id="c91fe-438">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c91fe-438">&lt;optional&gt;</span></span> | <span data-ttu-id="c91fe-439">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="c91fe-439">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="c91fe-440">Boolean</span><span class="sxs-lookup"><span data-stu-id="c91fe-440">Boolean</span></span> |  <span data-ttu-id="c91fe-441">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c91fe-441">&lt;optional&gt;</span></span> | <span data-ttu-id="c91fe-p125">提供されたトークンを Outlook REST API または Exchange Web サービスに使用するかどうかを決定します。既定値は、`false` です。</span><span class="sxs-lookup"><span data-stu-id="c91fe-p125">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="c91fe-444">Object</span><span class="sxs-lookup"><span data-stu-id="c91fe-444">Object</span></span> |  <span data-ttu-id="c91fe-445">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c91fe-445">&lt;optional&gt;</span></span> | <span data-ttu-id="c91fe-446">非同期メソッドに渡される状態データ。</span><span class="sxs-lookup"><span data-stu-id="c91fe-446">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="c91fe-447">function</span><span class="sxs-lookup"><span data-stu-id="c91fe-447">function</span></span>||<span data-ttu-id="c91fe-p126">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="c91fe-p126">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c91fe-450">要件</span><span class="sxs-lookup"><span data-stu-id="c91fe-450">Requirements</span></span>

|<span data-ttu-id="c91fe-451">要件</span><span class="sxs-lookup"><span data-stu-id="c91fe-451">Requirement</span></span>| <span data-ttu-id="c91fe-452">値</span><span class="sxs-lookup"><span data-stu-id="c91fe-452">Value</span></span>|
|---|---|
|[<span data-ttu-id="c91fe-453">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c91fe-453">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c91fe-454">1.5</span><span class="sxs-lookup"><span data-stu-id="c91fe-454">1.5</span></span> |
|[<span data-ttu-id="c91fe-455">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c91fe-455">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c91fe-456">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c91fe-456">ReadItem</span></span>|
|[<span data-ttu-id="c91fe-457">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c91fe-457">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c91fe-458">新規作成と閲覧</span><span class="sxs-lookup"><span data-stu-id="c91fe-458">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="c91fe-459">例</span><span class="sxs-lookup"><span data-stu-id="c91fe-459">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="c91fe-460">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="c91fe-460">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="c91fe-461">Exchange Server から添付ファイルやアイテムを取得するために使うトークンを含む文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="c91fe-461">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="c91fe-p127">`getCallbackTokenAsync` メソッドは、ユーザーのメールボックスをホストする Exchange Server から不透明なトークンを取得する非同期の呼び出しを行います。コールバック トークンの有効期間は 5 分です。</span><span class="sxs-lookup"><span data-stu-id="c91fe-p127">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="c91fe-p128">トークンと予定の識別子またはアイテムの識別子をサードパーティ システムに渡すことができます。サードパーティ システムは、トークンをベアラー承認トークンとして使用し、Exchange Web サービス (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) または [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) 操作を呼び出して、添付ファイルまたはアイテムを返します。たとえば、[選択したアイテムから添付ファイルを取得する](/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="c91fe-p128">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="c91fe-467">アプリが閲覧モードで `getCallbackTokenAsync` メソッドを呼び出すには、アプリのマニフェスト内に **ReadItem** アクセス許可が指定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="c91fe-467">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="c91fe-p129">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出してアイテムの識別子を `getCallbackTokenAsync` メソッドに渡す必要があります。アプリには、`saveAsync` メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="c91fe-p129">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c91fe-470">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c91fe-470">Parameters</span></span>

|<span data-ttu-id="c91fe-471">名前</span><span class="sxs-lookup"><span data-stu-id="c91fe-471">Name</span></span>| <span data-ttu-id="c91fe-472">型</span><span class="sxs-lookup"><span data-stu-id="c91fe-472">Type</span></span>| <span data-ttu-id="c91fe-473">属性</span><span class="sxs-lookup"><span data-stu-id="c91fe-473">Attributes</span></span>| <span data-ttu-id="c91fe-474">説明</span><span class="sxs-lookup"><span data-stu-id="c91fe-474">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="c91fe-475">function</span><span class="sxs-lookup"><span data-stu-id="c91fe-475">function</span></span>||<span data-ttu-id="c91fe-p130">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="c91fe-p130">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="c91fe-478">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="c91fe-478">Object</span></span>| <span data-ttu-id="c91fe-479">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c91fe-479">&lt;optional&gt;</span></span>|<span data-ttu-id="c91fe-480">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="c91fe-480">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c91fe-481">要件</span><span class="sxs-lookup"><span data-stu-id="c91fe-481">Requirements</span></span>

|<span data-ttu-id="c91fe-482">要件</span><span class="sxs-lookup"><span data-stu-id="c91fe-482">Requirement</span></span>| <span data-ttu-id="c91fe-483">値</span><span class="sxs-lookup"><span data-stu-id="c91fe-483">Value</span></span>|
|---|---|
|[<span data-ttu-id="c91fe-484">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c91fe-484">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c91fe-485">1.3</span><span class="sxs-lookup"><span data-stu-id="c91fe-485">1.3</span></span>|
|[<span data-ttu-id="c91fe-486">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c91fe-486">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c91fe-487">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c91fe-487">ReadItem</span></span>|
|[<span data-ttu-id="c91fe-488">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c91fe-488">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c91fe-489">新規作成と閲覧</span><span class="sxs-lookup"><span data-stu-id="c91fe-489">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="c91fe-490">例</span><span class="sxs-lookup"><span data-stu-id="c91fe-490">Example</span></span>

```javascript
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="c91fe-491">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="c91fe-491">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="c91fe-492">ユーザーと Office アドインを識別するトークンを取得します。</span><span class="sxs-lookup"><span data-stu-id="c91fe-492">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="c91fe-493">`getUserIdentityTokenAsync` メソッドは、[アドインとユーザーをサード パーティのシステムで識別して認証](/outlook/add-ins/authentication)することのできるトークンを返します。</span><span class="sxs-lookup"><span data-stu-id="c91fe-493">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="c91fe-494">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c91fe-494">Parameters</span></span>

|<span data-ttu-id="c91fe-495">名前</span><span class="sxs-lookup"><span data-stu-id="c91fe-495">Name</span></span>| <span data-ttu-id="c91fe-496">型</span><span class="sxs-lookup"><span data-stu-id="c91fe-496">Type</span></span>| <span data-ttu-id="c91fe-497">属性</span><span class="sxs-lookup"><span data-stu-id="c91fe-497">Attributes</span></span>| <span data-ttu-id="c91fe-498">説明</span><span class="sxs-lookup"><span data-stu-id="c91fe-498">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="c91fe-499">function</span><span class="sxs-lookup"><span data-stu-id="c91fe-499">function</span></span>||<span data-ttu-id="c91fe-500">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c91fe-500">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="c91fe-501">トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="c91fe-501">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="c91fe-502">Object</span><span class="sxs-lookup"><span data-stu-id="c91fe-502">Object</span></span>| <span data-ttu-id="c91fe-503">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c91fe-503">&lt;optional&gt;</span></span>|<span data-ttu-id="c91fe-504">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="c91fe-504">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c91fe-505">要件</span><span class="sxs-lookup"><span data-stu-id="c91fe-505">Requirements</span></span>

|<span data-ttu-id="c91fe-506">要件</span><span class="sxs-lookup"><span data-stu-id="c91fe-506">Requirement</span></span>| <span data-ttu-id="c91fe-507">値</span><span class="sxs-lookup"><span data-stu-id="c91fe-507">Value</span></span>|
|---|---|
|[<span data-ttu-id="c91fe-508">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c91fe-508">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c91fe-509">1.0</span><span class="sxs-lookup"><span data-stu-id="c91fe-509">1.0</span></span>|
|[<span data-ttu-id="c91fe-510">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c91fe-510">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c91fe-511">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c91fe-511">ReadItem</span></span>|
|[<span data-ttu-id="c91fe-512">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c91fe-512">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c91fe-513">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c91fe-513">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c91fe-514">例</span><span class="sxs-lookup"><span data-stu-id="c91fe-514">Example</span></span>

```javascript
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="c91fe-515">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="c91fe-515">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="c91fe-516">ユーザーのメールボックスをホストしている Exchange サーバー上の Exchange Web サービス (EWS) のサービスに対して非同期の要求を行います。</span><span class="sxs-lookup"><span data-stu-id="c91fe-516">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="c91fe-517">このメソッドは、次のシナリオではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c91fe-517">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="c91fe-518">Outlook on iOS または Android の場合</span><span class="sxs-lookup"><span data-stu-id="c91fe-518">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="c91fe-519">アドインが Gmail のメールボックスに読み込まれる場合</span><span class="sxs-lookup"><span data-stu-id="c91fe-519">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="c91fe-520">このような場合は、アドインでは [REST API を使用](/outlook/add-ins/use-rest-api)して、代わりにユーザーのメールボックスにアクセスする必要があります。</span><span class="sxs-lookup"><span data-stu-id="c91fe-520">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="c91fe-p131">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange. See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span><span class="sxs-lookup"><span data-stu-id="c91fe-p131">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange. See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="c91fe-523">`makeEwsRequestAsync` メソッドでは、フォルダー関連アイテムを要求できません。</span><span class="sxs-lookup"><span data-stu-id="c91fe-523">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="c91fe-524">XML 要求では UTF-8 エンコードを指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c91fe-524">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="c91fe-p132">`makeEwsRequestAsync` メソッドを使用するには、アドインに **ReadWriteMailbox** アクセス許可が必要です。**ReadWriteMailbox** アクセス許可と、`makeEwsRequestAsync` メソッドで呼び出せる EWS 操作の使い方については、「[ユーザーのメールボックスへのメール アドイン アクセスのアクセス許可を指定する](/outlook/add-ins/understanding-outlook-add-in-permissions)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c91fe-p132">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="c91fe-527">サーバー管理者は、クライアント アクセス サーバーの EWS ディレクトリで `OAuthAuthentication` を true に設定して、`makeEwsRequestAsync` メソッドで EWS 要求を行うことができるようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="c91fe-527">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="c91fe-528">バージョンの相違点</span><span class="sxs-lookup"><span data-stu-id="c91fe-528">Version differences</span></span>

<span data-ttu-id="c91fe-529">バージョン 15.0.4535.1004 より前のバージョンの Outlook で実行しているメール アプリで `makeEwsRequestAsync` メソッドを使う場合は、エンコード値を `ISO-8859-1` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c91fe-529">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="c91fe-p133">Outlook on the web でメール アプリを実行している場合は、エンコード値を設定する必要はありません。mailbox.diagnostics.hostName プロパティを使って、メール アプリを Outlook で実行しているのか、Outlook on the web で実行しているのかを確認できます。mailbox.diagnostics.hostVersion プロパティを使って、どのバージョンの Outlook を使って実行しているかを確認できます。</span><span class="sxs-lookup"><span data-stu-id="c91fe-p133">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c91fe-533">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c91fe-533">Parameters</span></span>

|<span data-ttu-id="c91fe-534">名前</span><span class="sxs-lookup"><span data-stu-id="c91fe-534">Name</span></span>| <span data-ttu-id="c91fe-535">型</span><span class="sxs-lookup"><span data-stu-id="c91fe-535">Type</span></span>| <span data-ttu-id="c91fe-536">属性</span><span class="sxs-lookup"><span data-stu-id="c91fe-536">Attributes</span></span>| <span data-ttu-id="c91fe-537">説明</span><span class="sxs-lookup"><span data-stu-id="c91fe-537">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="c91fe-538">String</span><span class="sxs-lookup"><span data-stu-id="c91fe-538">String</span></span>||<span data-ttu-id="c91fe-539">EWS 要求です。</span><span class="sxs-lookup"><span data-stu-id="c91fe-539">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="c91fe-540">function</span><span class="sxs-lookup"><span data-stu-id="c91fe-540">function</span></span>||<span data-ttu-id="c91fe-541">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c91fe-541">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="c91fe-p134">The XML result of the EWS call is provided as a string in the `asyncResult.value` property. If the result exceeds 1 MB in size, an error message is returned instead.</span><span class="sxs-lookup"><span data-stu-id="c91fe-p134">The XML result of the EWS call is provided as a string in the `asyncResult.value` property. If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="c91fe-544">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="c91fe-544">Object</span></span>| <span data-ttu-id="c91fe-545">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c91fe-545">&lt;optional&gt;</span></span>|<span data-ttu-id="c91fe-546">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="c91fe-546">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c91fe-547">要件</span><span class="sxs-lookup"><span data-stu-id="c91fe-547">Requirements</span></span>

|<span data-ttu-id="c91fe-548">要件</span><span class="sxs-lookup"><span data-stu-id="c91fe-548">Requirement</span></span>| <span data-ttu-id="c91fe-549">値</span><span class="sxs-lookup"><span data-stu-id="c91fe-549">Value</span></span>|
|---|---|
|[<span data-ttu-id="c91fe-550">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c91fe-550">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c91fe-551">1.0</span><span class="sxs-lookup"><span data-stu-id="c91fe-551">1.0</span></span>|
|[<span data-ttu-id="c91fe-552">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c91fe-552">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c91fe-553">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="c91fe-553">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="c91fe-554">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c91fe-554">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c91fe-555">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c91fe-555">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c91fe-556">例</span><span class="sxs-lookup"><span data-stu-id="c91fe-556">Example</span></span>

<span data-ttu-id="c91fe-557">次の例は、`GetItem` 操作を使ってアイテムの件名を取得するため、`makeEwsRequestAsync` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="c91fe-557">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="c91fe-558">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c91fe-558">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="c91fe-559">サポートされているイベントの種類のイベント ハンドラーを削除します。</span><span class="sxs-lookup"><span data-stu-id="c91fe-559">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="c91fe-560">現在、サポートされているイベントの種類は `Office.EventType.ItemChanged` だけです。</span><span class="sxs-lookup"><span data-stu-id="c91fe-560">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c91fe-561">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c91fe-561">Parameters</span></span>

| <span data-ttu-id="c91fe-562">名前</span><span class="sxs-lookup"><span data-stu-id="c91fe-562">Name</span></span> | <span data-ttu-id="c91fe-563">型</span><span class="sxs-lookup"><span data-stu-id="c91fe-563">Type</span></span> | <span data-ttu-id="c91fe-564">属性</span><span class="sxs-lookup"><span data-stu-id="c91fe-564">Attributes</span></span> | <span data-ttu-id="c91fe-565">説明</span><span class="sxs-lookup"><span data-stu-id="c91fe-565">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="c91fe-566">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="c91fe-566">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="c91fe-567">ハンドラーを取り消すイベント。</span><span class="sxs-lookup"><span data-stu-id="c91fe-567">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="c91fe-568">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="c91fe-568">Object</span></span> | <span data-ttu-id="c91fe-569">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c91fe-569">&lt;optional&gt;</span></span> | <span data-ttu-id="c91fe-570">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="c91fe-570">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="c91fe-571">Object</span><span class="sxs-lookup"><span data-stu-id="c91fe-571">Object</span></span> | <span data-ttu-id="c91fe-572">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c91fe-572">&lt;optional&gt;</span></span> | <span data-ttu-id="c91fe-573">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="c91fe-573">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="c91fe-574">function</span><span class="sxs-lookup"><span data-stu-id="c91fe-574">function</span></span>| <span data-ttu-id="c91fe-575">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c91fe-575">&lt;optional&gt;</span></span>|<span data-ttu-id="c91fe-576">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c91fe-576">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c91fe-577">要件</span><span class="sxs-lookup"><span data-stu-id="c91fe-577">Requirements</span></span>

|<span data-ttu-id="c91fe-578">要件</span><span class="sxs-lookup"><span data-stu-id="c91fe-578">Requirement</span></span>| <span data-ttu-id="c91fe-579">値</span><span class="sxs-lookup"><span data-stu-id="c91fe-579">Value</span></span>|
|---|---|
|[<span data-ttu-id="c91fe-580">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c91fe-580">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c91fe-581">1.5</span><span class="sxs-lookup"><span data-stu-id="c91fe-581">1.5</span></span> |
|[<span data-ttu-id="c91fe-582">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c91fe-582">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c91fe-583">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c91fe-583">ReadItem</span></span> |
|[<span data-ttu-id="c91fe-584">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c91fe-584">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c91fe-585">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c91fe-585">Compose or Read</span></span>|
