---
title: Office. メールボックス要件セット1.6
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 9b91a61d301434886723a55eca9608f004f598eb
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871935"
---
# <a name="mailbox"></a><span data-ttu-id="0ba03-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="0ba03-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="0ba03-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="0ba03-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="0ba03-104">Microsoft Outlook と Microsoft Outlook on the web の Outlook アドイン オブジェクト モデルへのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="0ba03-104">Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="0ba03-105">要件</span><span class="sxs-lookup"><span data-stu-id="0ba03-105">Requirements</span></span>

|<span data-ttu-id="0ba03-106">要件</span><span class="sxs-lookup"><span data-stu-id="0ba03-106">Requirement</span></span>| <span data-ttu-id="0ba03-107">値</span><span class="sxs-lookup"><span data-stu-id="0ba03-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="0ba03-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0ba03-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0ba03-109">1.0</span><span class="sxs-lookup"><span data-stu-id="0ba03-109">1.0</span></span>|
|[<span data-ttu-id="0ba03-110">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0ba03-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0ba03-111">制限あり</span><span class="sxs-lookup"><span data-stu-id="0ba03-111">Restricted</span></span>|
|[<span data-ttu-id="0ba03-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0ba03-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0ba03-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0ba03-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="0ba03-114">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="0ba03-114">Members and methods</span></span>

| <span data-ttu-id="0ba03-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="0ba03-115">Member</span></span> | <span data-ttu-id="0ba03-116">種類</span><span class="sxs-lookup"><span data-stu-id="0ba03-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="0ba03-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="0ba03-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="0ba03-118">メンバー</span><span class="sxs-lookup"><span data-stu-id="0ba03-118">Member</span></span> |
| [<span data-ttu-id="0ba03-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="0ba03-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="0ba03-120">メンバー</span><span class="sxs-lookup"><span data-stu-id="0ba03-120">Member</span></span> |
| [<span data-ttu-id="0ba03-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="0ba03-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="0ba03-122">メソッド</span><span class="sxs-lookup"><span data-stu-id="0ba03-122">Method</span></span> |
| [<span data-ttu-id="0ba03-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="0ba03-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="0ba03-124">メソッド</span><span class="sxs-lookup"><span data-stu-id="0ba03-124">Method</span></span> |
| [<span data-ttu-id="0ba03-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="0ba03-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="0ba03-126">メソッド</span><span class="sxs-lookup"><span data-stu-id="0ba03-126">Method</span></span> |
| [<span data-ttu-id="0ba03-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="0ba03-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="0ba03-128">メソッド</span><span class="sxs-lookup"><span data-stu-id="0ba03-128">Method</span></span> |
| [<span data-ttu-id="0ba03-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="0ba03-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="0ba03-130">メソッド</span><span class="sxs-lookup"><span data-stu-id="0ba03-130">Method</span></span> |
| [<span data-ttu-id="0ba03-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="0ba03-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="0ba03-132">メソッド</span><span class="sxs-lookup"><span data-stu-id="0ba03-132">Method</span></span> |
| [<span data-ttu-id="0ba03-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="0ba03-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="0ba03-134">メソッド</span><span class="sxs-lookup"><span data-stu-id="0ba03-134">Method</span></span> |
| [<span data-ttu-id="0ba03-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="0ba03-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="0ba03-136">メソッド</span><span class="sxs-lookup"><span data-stu-id="0ba03-136">Method</span></span> |
| [<span data-ttu-id="0ba03-137">displayNewMessageForm</span><span class="sxs-lookup"><span data-stu-id="0ba03-137">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="0ba03-138">メソッド</span><span class="sxs-lookup"><span data-stu-id="0ba03-138">Method</span></span> |
| [<span data-ttu-id="0ba03-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="0ba03-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="0ba03-140">メソッド</span><span class="sxs-lookup"><span data-stu-id="0ba03-140">Method</span></span> |
| [<span data-ttu-id="0ba03-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="0ba03-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="0ba03-142">メソッド</span><span class="sxs-lookup"><span data-stu-id="0ba03-142">Method</span></span> |
| [<span data-ttu-id="0ba03-143">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="0ba03-143">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="0ba03-144">メソッド</span><span class="sxs-lookup"><span data-stu-id="0ba03-144">Method</span></span> |
| [<span data-ttu-id="0ba03-145">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="0ba03-145">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="0ba03-146">メソッド</span><span class="sxs-lookup"><span data-stu-id="0ba03-146">Method</span></span> |
| [<span data-ttu-id="0ba03-147">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="0ba03-147">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="0ba03-148">メソッド</span><span class="sxs-lookup"><span data-stu-id="0ba03-148">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="0ba03-149">名前空間</span><span class="sxs-lookup"><span data-stu-id="0ba03-149">Namespaces</span></span>

<span data-ttu-id="0ba03-150">[diagnostics](Office.context.mailbox.diagnostics.md): Outlook アドインに診断情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="0ba03-150">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="0ba03-151">[item](Office.context.mailbox.item.md):Outlook アドインのメッセージや予定にアクセスするためのメソッドとプロパティを提供します。</span><span class="sxs-lookup"><span data-stu-id="0ba03-151">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="0ba03-152">[userProfile](Office.context.mailbox.userProfile.md):Outlook アドインのユーザーに関する情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="0ba03-152">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="0ba03-153">メンバー</span><span class="sxs-lookup"><span data-stu-id="0ba03-153">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="0ba03-154">ewsUrl: 文字列</span><span class="sxs-lookup"><span data-stu-id="0ba03-154">ewsUrl :String</span></span>

<span data-ttu-id="0ba03-p101">このメール アカウントの Exchange Web サービス (EWS) エンドポイントの URL を取得します。読み取りモードのみです。</span><span class="sxs-lookup"><span data-stu-id="0ba03-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="0ba03-157">このメンバーは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="0ba03-157">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="0ba03-p102">`ewsUrl` 値は、リモート サービスで、ユーザーのメールボックスに EWS 呼び出しを行うために使用することができます。たとえば、[選択した項目から添付ファイルを取得する](/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="0ba03-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="0ba03-160">閲覧モードで `ewsUrl` メンバーを呼び出すには、アプリでマニフェスト内に **ReadItem** アクセス許可を指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="0ba03-160">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="0ba03-p103">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback)メソッドを呼び出してから、`ewsUrl`メンバーを使用する必要があります。アプリには、`saveAsync`メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="0ba03-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="0ba03-163">型</span><span class="sxs-lookup"><span data-stu-id="0ba03-163">Type</span></span>

*   <span data-ttu-id="0ba03-164">String</span><span class="sxs-lookup"><span data-stu-id="0ba03-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0ba03-165">要件</span><span class="sxs-lookup"><span data-stu-id="0ba03-165">Requirements</span></span>

|<span data-ttu-id="0ba03-166">要件</span><span class="sxs-lookup"><span data-stu-id="0ba03-166">Requirement</span></span>| <span data-ttu-id="0ba03-167">値</span><span class="sxs-lookup"><span data-stu-id="0ba03-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="0ba03-168">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0ba03-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0ba03-169">1.0</span><span class="sxs-lookup"><span data-stu-id="0ba03-169">1.0</span></span>|
|[<span data-ttu-id="0ba03-170">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0ba03-170">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0ba03-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0ba03-171">ReadItem</span></span>|
|[<span data-ttu-id="0ba03-172">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0ba03-172">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0ba03-173">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0ba03-173">Compose or Read</span></span>|

#### <a name="resturl-string"></a><span data-ttu-id="0ba03-174">restUrl: 文字列</span><span class="sxs-lookup"><span data-stu-id="0ba03-174">restUrl :String</span></span>

<span data-ttu-id="0ba03-175">この電子メール アカウントの REST エンドポイントの URL を取得します。</span><span class="sxs-lookup"><span data-stu-id="0ba03-175">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="0ba03-176">`restUrl` 値は、ユーザーのメールボックスに [REST API](/outlook/rest/) 呼び出しを行うために使用することができます。</span><span class="sxs-lookup"><span data-stu-id="0ba03-176">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="0ba03-177">閲覧モードで `restUrl` メンバーを呼び出すには、アプリでマニフェスト内に **ReadItem** アクセス許可を指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="0ba03-177">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="0ba03-p104">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback)メソッドを呼び出してから、`restUrl`メンバーを使用する必要があります。アプリには、`saveAsync`メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="0ba03-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="0ba03-180">型</span><span class="sxs-lookup"><span data-stu-id="0ba03-180">Type</span></span>

*   <span data-ttu-id="0ba03-181">String</span><span class="sxs-lookup"><span data-stu-id="0ba03-181">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0ba03-182">要件</span><span class="sxs-lookup"><span data-stu-id="0ba03-182">Requirements</span></span>

|<span data-ttu-id="0ba03-183">要件</span><span class="sxs-lookup"><span data-stu-id="0ba03-183">Requirement</span></span>| <span data-ttu-id="0ba03-184">値</span><span class="sxs-lookup"><span data-stu-id="0ba03-184">Value</span></span>|
|---|---|
|[<span data-ttu-id="0ba03-185">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0ba03-185">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0ba03-186">1.5</span><span class="sxs-lookup"><span data-stu-id="0ba03-186">1.5</span></span> |
|[<span data-ttu-id="0ba03-187">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0ba03-187">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0ba03-188">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0ba03-188">ReadItem</span></span>|
|[<span data-ttu-id="0ba03-189">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0ba03-189">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0ba03-190">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0ba03-190">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="0ba03-191">メソッド</span><span class="sxs-lookup"><span data-stu-id="0ba03-191">Methods</span></span>

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="0ba03-192">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="0ba03-192">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="0ba03-193">サポートされているイベントのイベント ハンドラーを追加します。</span><span class="sxs-lookup"><span data-stu-id="0ba03-193">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="0ba03-194">現在サポートされている唯一のイベントの種類は `Office.EventType.ItemChanged` です。これはユーザーが新しいアイテムを選択したときに呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="0ba03-194">Currently the only supported event type is `Office.EventType.ItemChanged`, which is invoked when the user selects a new item.</span></span> <span data-ttu-id="0ba03-195">このイベントは、ピン留め可能な作業ウィンドウを実装するアドインで使用され、現在選択されているアイテムに基づいて作業ウィンドウ UI をアドインで更新できるようにします。</span><span class="sxs-lookup"><span data-stu-id="0ba03-195">This event is used by add-ins that implement a pinnable task pane, and allows the add-in to refresh the task pane UI based on the currently selected item.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0ba03-196">パラメーター</span><span class="sxs-lookup"><span data-stu-id="0ba03-196">Parameters</span></span>

| <span data-ttu-id="0ba03-197">名前</span><span class="sxs-lookup"><span data-stu-id="0ba03-197">Name</span></span> | <span data-ttu-id="0ba03-198">型</span><span class="sxs-lookup"><span data-stu-id="0ba03-198">Type</span></span> | <span data-ttu-id="0ba03-199">属性</span><span class="sxs-lookup"><span data-stu-id="0ba03-199">Attributes</span></span> | <span data-ttu-id="0ba03-200">説明</span><span class="sxs-lookup"><span data-stu-id="0ba03-200">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="0ba03-201">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="0ba03-201">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="0ba03-202">ハンドラを呼び出す必要のあるイベント。</span><span class="sxs-lookup"><span data-stu-id="0ba03-202">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="0ba03-203">関数</span><span class="sxs-lookup"><span data-stu-id="0ba03-203">Function</span></span> || <span data-ttu-id="0ba03-p106">イベントを処理する関数。この関数は、オブジェクト リテラルである単一パラメータを受け入れる必要があります。パラメータの `type` プロパティは、`addHandlerAsync` に渡される `eventType` パラメータと一致します。</span><span class="sxs-lookup"><span data-stu-id="0ba03-p106">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="0ba03-207">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="0ba03-207">Object</span></span> | <span data-ttu-id="0ba03-208">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="0ba03-208">&lt;optional&gt;</span></span> | <span data-ttu-id="0ba03-209">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="0ba03-209">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="0ba03-210">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="0ba03-210">Object</span></span> | <span data-ttu-id="0ba03-211">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0ba03-211">&lt;optional&gt;</span></span> | <span data-ttu-id="0ba03-212">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="0ba03-212">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="0ba03-213">function</span><span class="sxs-lookup"><span data-stu-id="0ba03-213">function</span></span>| <span data-ttu-id="0ba03-214">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0ba03-214">&lt;optional&gt;</span></span>|<span data-ttu-id="0ba03-215">メソッドが完了すると、`callback` パラメータに渡された関数が、単一パラメータ `asyncResult` で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="0ba03-215">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0ba03-216">要件</span><span class="sxs-lookup"><span data-stu-id="0ba03-216">Requirements</span></span>

|<span data-ttu-id="0ba03-217">要件</span><span class="sxs-lookup"><span data-stu-id="0ba03-217">Requirement</span></span>| <span data-ttu-id="0ba03-218">値</span><span class="sxs-lookup"><span data-stu-id="0ba03-218">Value</span></span>|
|---|---|
|[<span data-ttu-id="0ba03-219">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0ba03-219">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0ba03-220">1.5</span><span class="sxs-lookup"><span data-stu-id="0ba03-220">1.5</span></span> |
|[<span data-ttu-id="0ba03-221">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0ba03-221">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0ba03-222">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0ba03-222">ReadItem</span></span> |
|[<span data-ttu-id="0ba03-223">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0ba03-223">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0ba03-224">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0ba03-224">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0ba03-225">例</span><span class="sxs-lookup"><span data-stu-id="0ba03-225">Example</span></span>

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

####  <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="0ba03-226">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="0ba03-226">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="0ba03-227">REST 形式のアイテム ID を EWS 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="0ba03-227">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="0ba03-228">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="0ba03-228">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="0ba03-p107">REST API ([Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) や [Microsoft Graph](https://graph.microsoft.io/) など) 経由で取得された項目 ID は、Exchange Web サービス (EWS) で使用される形式とは異なる形式を使用します。`convertToEwsId` メソッドは、REST 形式の ID を EWS 用の適切な形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="0ba03-p107">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0ba03-231">パラメーター</span><span class="sxs-lookup"><span data-stu-id="0ba03-231">Parameters</span></span>

|<span data-ttu-id="0ba03-232">名前</span><span class="sxs-lookup"><span data-stu-id="0ba03-232">Name</span></span>| <span data-ttu-id="0ba03-233">種類</span><span class="sxs-lookup"><span data-stu-id="0ba03-233">Type</span></span>| <span data-ttu-id="0ba03-234">説明</span><span class="sxs-lookup"><span data-stu-id="0ba03-234">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="0ba03-235">String</span><span class="sxs-lookup"><span data-stu-id="0ba03-235">String</span></span>|<span data-ttu-id="0ba03-236">Outlook REST API 用に書式設定された項目 ID</span><span class="sxs-lookup"><span data-stu-id="0ba03-236">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="0ba03-237">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="0ba03-237">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.restversion)|<span data-ttu-id="0ba03-238">項目 ID の取得に使用された Outlook REST API のバージョンを示す値。</span><span class="sxs-lookup"><span data-stu-id="0ba03-238">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0ba03-239">要件</span><span class="sxs-lookup"><span data-stu-id="0ba03-239">Requirements</span></span>

|<span data-ttu-id="0ba03-240">要件</span><span class="sxs-lookup"><span data-stu-id="0ba03-240">Requirement</span></span>| <span data-ttu-id="0ba03-241">値</span><span class="sxs-lookup"><span data-stu-id="0ba03-241">Value</span></span>|
|---|---|
|[<span data-ttu-id="0ba03-242">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0ba03-242">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0ba03-243">1.3</span><span class="sxs-lookup"><span data-stu-id="0ba03-243">1.3</span></span>|
|[<span data-ttu-id="0ba03-244">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0ba03-244">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0ba03-245">制限あり</span><span class="sxs-lookup"><span data-stu-id="0ba03-245">Restricted</span></span>|
|[<span data-ttu-id="0ba03-246">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0ba03-246">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0ba03-247">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0ba03-247">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0ba03-248">戻り値:</span><span class="sxs-lookup"><span data-stu-id="0ba03-248">Returns:</span></span>

<span data-ttu-id="0ba03-249">種類: 文字列</span><span class="sxs-lookup"><span data-stu-id="0ba03-249">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="0ba03-250">例</span><span class="sxs-lookup"><span data-stu-id="0ba03-250">Example</span></span>

```javascript
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime"></a><span data-ttu-id="0ba03-251">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="0ba03-251">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)}</span></span>

<span data-ttu-id="0ba03-252">クライアントの現地時間の時間情報が含まれている辞書を取得します。</span><span class="sxs-lookup"><span data-stu-id="0ba03-252">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="0ba03-p108">Outlook 用メール アプリや Outlook Web App で使う日付と時刻は、異なるタイム ゾーンを使うことができます。Outlook では、クライアント コンピューターのタイム ゾーンを使います。Outlook Web App では、Exchange 管理センター (EAC) で設定されたタイム ゾーンを使います。ユーザー インターフェイスに表示される値が、常にユーザーが期待するタイム ゾーンと一致するように日付と時刻の値を処理する必要があります。</span><span class="sxs-lookup"><span data-stu-id="0ba03-p108">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="0ba03-p109">Outlook でメール アプリが実行されている場合、`convertToLocalClientTime` メソッドは、クライアント コンピューターのタイム ゾーンに設定された値で辞書オブジェクトを返します。Outlook Web Apps でメール アプリが実行されている場合、`convertToLocalClientTime` メソッドは、EAC で指定されたタイム ゾーンに設定された値で辞書オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="0ba03-p109">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0ba03-258">パラメーター</span><span class="sxs-lookup"><span data-stu-id="0ba03-258">Parameters</span></span>

|<span data-ttu-id="0ba03-259">名前</span><span class="sxs-lookup"><span data-stu-id="0ba03-259">Name</span></span>| <span data-ttu-id="0ba03-260">種類</span><span class="sxs-lookup"><span data-stu-id="0ba03-260">Type</span></span>| <span data-ttu-id="0ba03-261">説明</span><span class="sxs-lookup"><span data-stu-id="0ba03-261">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="0ba03-262">日付</span><span class="sxs-lookup"><span data-stu-id="0ba03-262">Date</span></span>|<span data-ttu-id="0ba03-263">Date オブジェクト</span><span class="sxs-lookup"><span data-stu-id="0ba03-263">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0ba03-264">要件</span><span class="sxs-lookup"><span data-stu-id="0ba03-264">Requirements</span></span>

|<span data-ttu-id="0ba03-265">要件</span><span class="sxs-lookup"><span data-stu-id="0ba03-265">Requirement</span></span>| <span data-ttu-id="0ba03-266">値</span><span class="sxs-lookup"><span data-stu-id="0ba03-266">Value</span></span>|
|---|---|
|[<span data-ttu-id="0ba03-267">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0ba03-267">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0ba03-268">1.0</span><span class="sxs-lookup"><span data-stu-id="0ba03-268">1.0</span></span>|
|[<span data-ttu-id="0ba03-269">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0ba03-269">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0ba03-270">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0ba03-270">ReadItem</span></span>|
|[<span data-ttu-id="0ba03-271">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0ba03-271">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0ba03-272">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0ba03-272">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0ba03-273">戻り値:</span><span class="sxs-lookup"><span data-stu-id="0ba03-273">Returns:</span></span>

<span data-ttu-id="0ba03-274">種類:[LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="0ba03-274">Type: [LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)</span></span>

####  <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="0ba03-275">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="0ba03-275">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="0ba03-276">EWS 形式のアイテム ID を REST 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="0ba03-276">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="0ba03-277">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="0ba03-277">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="0ba03-p110">EWS 経由または `itemId` プロパティ経由で取得される項目 ID では、REST API ([Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) または [Microsoft Graph](https://graph.microsoft.io/) など) で使用される形式とは異なる形式を使用します。`convertToRestId` メソッドは、EWS 形式の ID を REST 用の適切な形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="0ba03-p110">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0ba03-280">パラメーター</span><span class="sxs-lookup"><span data-stu-id="0ba03-280">Parameters</span></span>

|<span data-ttu-id="0ba03-281">名前</span><span class="sxs-lookup"><span data-stu-id="0ba03-281">Name</span></span>| <span data-ttu-id="0ba03-282">種類</span><span class="sxs-lookup"><span data-stu-id="0ba03-282">Type</span></span>| <span data-ttu-id="0ba03-283">説明</span><span class="sxs-lookup"><span data-stu-id="0ba03-283">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="0ba03-284">文字列</span><span class="sxs-lookup"><span data-stu-id="0ba03-284">String</span></span>|<span data-ttu-id="0ba03-285">Exchange Web サービス (EWS) 用に書式設定された項目 ID</span><span class="sxs-lookup"><span data-stu-id="0ba03-285">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="0ba03-286">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="0ba03-286">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.restversion)|<span data-ttu-id="0ba03-287">変換後の ID とともに使用される Outlook REST API のバージョンを示す値。</span><span class="sxs-lookup"><span data-stu-id="0ba03-287">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0ba03-288">要件</span><span class="sxs-lookup"><span data-stu-id="0ba03-288">Requirements</span></span>

|<span data-ttu-id="0ba03-289">要件</span><span class="sxs-lookup"><span data-stu-id="0ba03-289">Requirement</span></span>| <span data-ttu-id="0ba03-290">値</span><span class="sxs-lookup"><span data-stu-id="0ba03-290">Value</span></span>|
|---|---|
|[<span data-ttu-id="0ba03-291">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0ba03-291">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0ba03-292">1.3</span><span class="sxs-lookup"><span data-stu-id="0ba03-292">1.3</span></span>|
|[<span data-ttu-id="0ba03-293">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0ba03-293">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0ba03-294">制限あり</span><span class="sxs-lookup"><span data-stu-id="0ba03-294">Restricted</span></span>|
|[<span data-ttu-id="0ba03-295">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0ba03-295">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0ba03-296">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0ba03-296">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0ba03-297">戻り値:</span><span class="sxs-lookup"><span data-stu-id="0ba03-297">Returns:</span></span>

<span data-ttu-id="0ba03-298">種類: 文字列</span><span class="sxs-lookup"><span data-stu-id="0ba03-298">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="0ba03-299">例</span><span class="sxs-lookup"><span data-stu-id="0ba03-299">Example</span></span>

```javascript
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="0ba03-300">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="0ba03-300">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="0ba03-301">時間情報が含まれている辞書から Date オブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="0ba03-301">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="0ba03-302">`convertToUtcClientTime` メソッドは、ローカルの日付と時刻が含まれる辞書を、ローカルの日付と時刻の正しい値をもつ Date オブジェクトに変換します。</span><span class="sxs-lookup"><span data-stu-id="0ba03-302">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0ba03-303">パラメーター</span><span class="sxs-lookup"><span data-stu-id="0ba03-303">Parameters</span></span>

|<span data-ttu-id="0ba03-304">名前</span><span class="sxs-lookup"><span data-stu-id="0ba03-304">Name</span></span>| <span data-ttu-id="0ba03-305">種類</span><span class="sxs-lookup"><span data-stu-id="0ba03-305">Type</span></span>| <span data-ttu-id="0ba03-306">説明</span><span class="sxs-lookup"><span data-stu-id="0ba03-306">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="0ba03-307">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="0ba03-307">LocalClientTime</span></span>](/javascript/api/outlook_1_6/office.LocalClientTime)|<span data-ttu-id="0ba03-308">変換するローカル時刻の値。</span><span class="sxs-lookup"><span data-stu-id="0ba03-308">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0ba03-309">要件</span><span class="sxs-lookup"><span data-stu-id="0ba03-309">Requirements</span></span>

|<span data-ttu-id="0ba03-310">要件</span><span class="sxs-lookup"><span data-stu-id="0ba03-310">Requirement</span></span>| <span data-ttu-id="0ba03-311">値</span><span class="sxs-lookup"><span data-stu-id="0ba03-311">Value</span></span>|
|---|---|
|[<span data-ttu-id="0ba03-312">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0ba03-312">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0ba03-313">1.0</span><span class="sxs-lookup"><span data-stu-id="0ba03-313">1.0</span></span>|
|[<span data-ttu-id="0ba03-314">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0ba03-314">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0ba03-315">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0ba03-315">ReadItem</span></span>|
|[<span data-ttu-id="0ba03-316">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0ba03-316">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0ba03-317">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0ba03-317">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0ba03-318">戻り値 :</span><span class="sxs-lookup"><span data-stu-id="0ba03-318">Returns:</span></span>

<span data-ttu-id="0ba03-319">時間が UTC で表現された Date オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="0ba03-319">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="0ba03-320">

<dt>種類</dt>

</span><span class="sxs-lookup"><span data-stu-id="0ba03-320">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="0ba03-321">Date</span><span class="sxs-lookup"><span data-stu-id="0ba03-321">Date</span></span></dd>

</dl>

####  <a name="displayappointmentformitemid"></a><span data-ttu-id="0ba03-322">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="0ba03-322">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="0ba03-323">既存の予定を表示します。</span><span class="sxs-lookup"><span data-stu-id="0ba03-323">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="0ba03-324">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="0ba03-324">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="0ba03-325">`displayAppointmentForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存の予定を開きます。</span><span class="sxs-lookup"><span data-stu-id="0ba03-325">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="0ba03-p111">Outlook for Mac では、このメソッドを使用して、定期的に繰り返される予定の一部ではない単発の予定、または定期的に繰り替えされる予定の元の予定を表示できます。ただし、一連の予定のインスタンスは表示できません。これは、Outlook for Mac では、定期的に繰り返されるインスタンスのプロパティ  (項目 ID を含む) にアクセスできないためです。</span><span class="sxs-lookup"><span data-stu-id="0ba03-p111">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="0ba03-328">Outlook Web App では、このメソッドで、指定されたフォームの本文が 32 KB 以下の文字数の場合に、そのフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="0ba03-328">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="0ba03-329">指定した項目識別子が既存の予定を識別しない場合には、クライアント コンピュータまたはデバイスで空白のウィンドウが開き、エラー メッセージは返されません。</span><span class="sxs-lookup"><span data-stu-id="0ba03-329">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0ba03-330">パラメーター</span><span class="sxs-lookup"><span data-stu-id="0ba03-330">Parameters</span></span>

|<span data-ttu-id="0ba03-331">名前</span><span class="sxs-lookup"><span data-stu-id="0ba03-331">Name</span></span>| <span data-ttu-id="0ba03-332">型</span><span class="sxs-lookup"><span data-stu-id="0ba03-332">Type</span></span>| <span data-ttu-id="0ba03-333">説明</span><span class="sxs-lookup"><span data-stu-id="0ba03-333">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="0ba03-334">String</span><span class="sxs-lookup"><span data-stu-id="0ba03-334">String</span></span>|<span data-ttu-id="0ba03-335">既存の予定表の予定の Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="0ba03-335">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0ba03-336">要件</span><span class="sxs-lookup"><span data-stu-id="0ba03-336">Requirements</span></span>

|<span data-ttu-id="0ba03-337">要件</span><span class="sxs-lookup"><span data-stu-id="0ba03-337">Requirement</span></span>| <span data-ttu-id="0ba03-338">値</span><span class="sxs-lookup"><span data-stu-id="0ba03-338">Value</span></span>|
|---|---|
|[<span data-ttu-id="0ba03-339">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0ba03-339">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0ba03-340">1.0</span><span class="sxs-lookup"><span data-stu-id="0ba03-340">1.0</span></span>|
|[<span data-ttu-id="0ba03-341">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0ba03-341">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0ba03-342">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0ba03-342">ReadItem</span></span>|
|[<span data-ttu-id="0ba03-343">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0ba03-343">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0ba03-344">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0ba03-344">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0ba03-345">例</span><span class="sxs-lookup"><span data-stu-id="0ba03-345">Example</span></span>

```javascript
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  <a name="displaymessageformitemid"></a><span data-ttu-id="0ba03-346">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="0ba03-346">displayMessageForm(itemId)</span></span>

<span data-ttu-id="0ba03-347">既存のメッセージを表示します。</span><span class="sxs-lookup"><span data-stu-id="0ba03-347">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="0ba03-348">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="0ba03-348">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="0ba03-349">`displayMessageForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存のメッセージを開きます。</span><span class="sxs-lookup"><span data-stu-id="0ba03-349">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="0ba03-350">Outlook Web App では、このメソッドは、指定されたフォームの本文が 32 KB 以下の文字数の場合に、そのフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="0ba03-350">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="0ba03-351">指定した項目識別子が既存のメッセージを識別しない場合には、クラアント コンピュータでメッセージは表示されず、エラー メッセージは返されません。</span><span class="sxs-lookup"><span data-stu-id="0ba03-351">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="0ba03-p112">予定を表す `itemId` が含まれる `displayMessageForm` を使用しないでください。`displayAppointmentForm` メソッドを使用して既存の予定を表示し、`displayNewAppointmentForm` を使用して新しい予定を作成するフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="0ba03-p112">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0ba03-354">パラメーター</span><span class="sxs-lookup"><span data-stu-id="0ba03-354">Parameters</span></span>

|<span data-ttu-id="0ba03-355">名前</span><span class="sxs-lookup"><span data-stu-id="0ba03-355">Name</span></span>| <span data-ttu-id="0ba03-356">型</span><span class="sxs-lookup"><span data-stu-id="0ba03-356">Type</span></span>| <span data-ttu-id="0ba03-357">説明</span><span class="sxs-lookup"><span data-stu-id="0ba03-357">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="0ba03-358">String</span><span class="sxs-lookup"><span data-stu-id="0ba03-358">String</span></span>|<span data-ttu-id="0ba03-359">既存のメッセージの Exchange Web サービス(EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="0ba03-359">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0ba03-360">要件</span><span class="sxs-lookup"><span data-stu-id="0ba03-360">Requirements</span></span>

|<span data-ttu-id="0ba03-361">要件</span><span class="sxs-lookup"><span data-stu-id="0ba03-361">Requirement</span></span>| <span data-ttu-id="0ba03-362">値</span><span class="sxs-lookup"><span data-stu-id="0ba03-362">Value</span></span>|
|---|---|
|[<span data-ttu-id="0ba03-363">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0ba03-363">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0ba03-364">1.0</span><span class="sxs-lookup"><span data-stu-id="0ba03-364">1.0</span></span>|
|[<span data-ttu-id="0ba03-365">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0ba03-365">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0ba03-366">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0ba03-366">ReadItem</span></span>|
|[<span data-ttu-id="0ba03-367">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0ba03-367">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0ba03-368">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0ba03-368">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0ba03-369">例</span><span class="sxs-lookup"><span data-stu-id="0ba03-369">Example</span></span>

```javascript
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="0ba03-370">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="0ba03-370">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="0ba03-371">新しい予定を作成するためのフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="0ba03-371">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="0ba03-372">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="0ba03-372">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="0ba03-p113">`displayNewAppointmentForm` メソッドを使用すると、ユーザーが新しい予定または会議を作成できるフォームが開きます。パラメータを指定すると、予定のフォーム フィールドにパラメータの内容が自動的に入力されます。</span><span class="sxs-lookup"><span data-stu-id="0ba03-p113">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="0ba03-p114">Outlook Web App および OWA for Devices では、このメソッドは出席者フィールドが含まれるフォームを常に表示します。入力因数として出席者を指定しない場合には、このメソッドは [**保存**] ボタンのあるフォームを表示します。出席者を指定した場合には、フォームにはその出席者と [**送信**] ボタンが含まれます。</span><span class="sxs-lookup"><span data-stu-id="0ba03-p114">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="0ba03-p115">Outlook リッチ クライアントおよび Outlook RT では、`requiredAttendees`、`optionalAttendees` または `resources` パラメータに出席者またはリソースを指定した場合、このメソッドは [**送信**] ボタンがある会議フォームを表示します。受信者を指定しない場合には、このメソッドは [**保存して閉じる**] ボタンがある予定フォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="0ba03-p115">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="0ba03-380">パラメーターのいずれかが指定のサイズ制限を超える場合、または不明なパラメーター名が指定されている場合は、例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="0ba03-380">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0ba03-381">パラメーター</span><span class="sxs-lookup"><span data-stu-id="0ba03-381">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="0ba03-382">すべてのパラメーターは省略可能です。</span><span class="sxs-lookup"><span data-stu-id="0ba03-382">All parameters are optional.</span></span>

|<span data-ttu-id="0ba03-383">名前</span><span class="sxs-lookup"><span data-stu-id="0ba03-383">Name</span></span>| <span data-ttu-id="0ba03-384">型</span><span class="sxs-lookup"><span data-stu-id="0ba03-384">Type</span></span>| <span data-ttu-id="0ba03-385">説明</span><span class="sxs-lookup"><span data-stu-id="0ba03-385">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="0ba03-386">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="0ba03-386">Object</span></span> | <span data-ttu-id="0ba03-387">新しい予定を記述するパラメータの辞書。</span><span class="sxs-lookup"><span data-stu-id="0ba03-387">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="0ba03-388">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="0ba03-388">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="0ba03-p116">予定への各必須出席者の電子メール アドレスを含む文字列の配列、または出席者の `EmailAddressDetails` オブジェクトを含む配列。この配列は、最大 100 エントリまでに制限されています。</span><span class="sxs-lookup"><span data-stu-id="0ba03-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="0ba03-391">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="0ba03-391">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="0ba03-p117">予定への各任意出席者の電子メール アドレスを含む文字列の配列、または出席者の `EmailAddressDetails` オブジェクトを含む配列。この配列は、最大 100 エントリまでに制限されています。</span><span class="sxs-lookup"><span data-stu-id="0ba03-p117">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="0ba03-394">日付</span><span class="sxs-lookup"><span data-stu-id="0ba03-394">Date</span></span> | <span data-ttu-id="0ba03-395">予定の開始日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="0ba03-395">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="0ba03-396">日付</span><span class="sxs-lookup"><span data-stu-id="0ba03-396">Date</span></span> | <span data-ttu-id="0ba03-397">予定の終了日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="0ba03-397">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="0ba03-398">String</span><span class="sxs-lookup"><span data-stu-id="0ba03-398">String</span></span> | <span data-ttu-id="0ba03-p118">予定の場所を含む文字列。文字列は最大 255 文字までに制限されています。</span><span class="sxs-lookup"><span data-stu-id="0ba03-p118">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="0ba03-401">配列。&lt; 文字列&gt;</span><span class="sxs-lookup"><span data-stu-id="0ba03-401">Array.&lt;String&gt;</span></span> | <span data-ttu-id="0ba03-p119">予定に必要なリソースを含む文字列の配列。配列は最大 100 エントリまでに制限されています。</span><span class="sxs-lookup"><span data-stu-id="0ba03-p119">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="0ba03-404">String</span><span class="sxs-lookup"><span data-stu-id="0ba03-404">String</span></span> | <span data-ttu-id="0ba03-p120">予定の件名を含む文字列。文字列は最大 255 文字までに制限されています。</span><span class="sxs-lookup"><span data-stu-id="0ba03-p120">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="0ba03-407">String</span><span class="sxs-lookup"><span data-stu-id="0ba03-407">String</span></span> | <span data-ttu-id="0ba03-p121">予定の本文。本文の内容は、最大サイズが 32 KB までに制限されています。</span><span class="sxs-lookup"><span data-stu-id="0ba03-p121">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="0ba03-410">要件</span><span class="sxs-lookup"><span data-stu-id="0ba03-410">Requirements</span></span>

|<span data-ttu-id="0ba03-411">要件</span><span class="sxs-lookup"><span data-stu-id="0ba03-411">Requirement</span></span>| <span data-ttu-id="0ba03-412">値</span><span class="sxs-lookup"><span data-stu-id="0ba03-412">Value</span></span>|
|---|---|
|[<span data-ttu-id="0ba03-413">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0ba03-413">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0ba03-414">1.0</span><span class="sxs-lookup"><span data-stu-id="0ba03-414">1.0</span></span>|
|[<span data-ttu-id="0ba03-415">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0ba03-415">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0ba03-416">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0ba03-416">ReadItem</span></span>|
|[<span data-ttu-id="0ba03-417">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0ba03-417">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0ba03-418">読み取り</span><span class="sxs-lookup"><span data-stu-id="0ba03-418">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0ba03-419">例</span><span class="sxs-lookup"><span data-stu-id="0ba03-419">Example</span></span>

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

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="0ba03-420">displayNewMessageForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="0ba03-420">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="0ba03-421">新しいメッセージを作成するためのフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="0ba03-421">Displays a form for creating a new message.</span></span>

<span data-ttu-id="0ba03-p122">The `displayNewMessageForm` method opens a form that enables the user to create a new message. If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span><span class="sxs-lookup"><span data-stu-id="0ba03-p122">The `displayNewMessageForm` method opens a form that enables the user to create a new message. If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="0ba03-424">パラメータのいずれかが指定されたサイズ制限を超えた場合、または不明なパラメータ名が指定された場合には、例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="0ba03-424">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0ba03-425">パラメーター</span><span class="sxs-lookup"><span data-stu-id="0ba03-425">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="0ba03-426">すべてのパラメーターは省略可能です。</span><span class="sxs-lookup"><span data-stu-id="0ba03-426">All parameters are optional.</span></span>

|<span data-ttu-id="0ba03-427">名前</span><span class="sxs-lookup"><span data-stu-id="0ba03-427">Name</span></span>| <span data-ttu-id="0ba03-428">型</span><span class="sxs-lookup"><span data-stu-id="0ba03-428">Type</span></span>| <span data-ttu-id="0ba03-429">説明</span><span class="sxs-lookup"><span data-stu-id="0ba03-429">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="0ba03-430">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="0ba03-430">Object</span></span> | <span data-ttu-id="0ba03-431">新しいメッセージを記述するパラメータの辞書。</span><span class="sxs-lookup"><span data-stu-id="0ba03-431">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="0ba03-432">配列。&lt;文字列&gt; | 配列。&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="0ba03-432">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="0ba03-p123">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line. The array is limited to a maximum of 100 entries.</span><span class="sxs-lookup"><span data-stu-id="0ba03-p123">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="0ba03-435">配列。&lt;文字列&gt; | 配列。&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="0ba03-435">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="0ba03-p124">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line. The array is limited to a maximum of 100 entries.</span><span class="sxs-lookup"><span data-stu-id="0ba03-p124">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="0ba03-438">配列。&lt;文字列&gt; | 配列。&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="0ba03-438">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="0ba03-p125">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line. The array is limited to a maximum of 100 entries.</span><span class="sxs-lookup"><span data-stu-id="0ba03-p125">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="0ba03-441">String</span><span class="sxs-lookup"><span data-stu-id="0ba03-441">String</span></span> | <span data-ttu-id="0ba03-p126">A string containing the subject of the message. The string is limited to a maximum of 255 characters.</span><span class="sxs-lookup"><span data-stu-id="0ba03-p126">A string containing the subject of the message. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="0ba03-444">String</span><span class="sxs-lookup"><span data-stu-id="0ba03-444">String</span></span> | <span data-ttu-id="0ba03-p127">The HTML body of the message. The body content is limited to a maximum size of 32 KB.</span><span class="sxs-lookup"><span data-stu-id="0ba03-p127">The HTML body of the message. The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="0ba03-447">配列。&lt;オブジェクト&gt;</span><span class="sxs-lookup"><span data-stu-id="0ba03-447">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="0ba03-448">添付ファイルまたは添付アイテムである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="0ba03-448">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="0ba03-449">String</span><span class="sxs-lookup"><span data-stu-id="0ba03-449">String</span></span> | <span data-ttu-id="0ba03-p128">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="0ba03-p128">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="0ba03-452">String</span><span class="sxs-lookup"><span data-stu-id="0ba03-452">String</span></span> | <span data-ttu-id="0ba03-453">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="0ba03-453">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="0ba03-454">文字列</span><span class="sxs-lookup"><span data-stu-id="0ba03-454">String</span></span> | <span data-ttu-id="0ba03-p129">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="0ba03-p129">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="0ba03-457">ブール値</span><span class="sxs-lookup"><span data-stu-id="0ba03-457">Boolean</span></span> | <span data-ttu-id="0ba03-p130">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="0ba03-p130">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="0ba03-460">String</span><span class="sxs-lookup"><span data-stu-id="0ba03-460">String</span></span> | <span data-ttu-id="0ba03-p131">Only used if `type` is set to `item`. The EWS item id of the existing e-mail you want to attach to the new message. This is a string up to 100 characters.</span><span class="sxs-lookup"><span data-stu-id="0ba03-p131">Only used if `type` is set to `item`. The EWS item id of the existing e-mail you want to attach to the new message. This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="0ba03-464">要件</span><span class="sxs-lookup"><span data-stu-id="0ba03-464">Requirements</span></span>

|<span data-ttu-id="0ba03-465">要件</span><span class="sxs-lookup"><span data-stu-id="0ba03-465">Requirement</span></span>| <span data-ttu-id="0ba03-466">値</span><span class="sxs-lookup"><span data-stu-id="0ba03-466">Value</span></span>|
|---|---|
|[<span data-ttu-id="0ba03-467">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0ba03-467">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0ba03-468">1.6</span><span class="sxs-lookup"><span data-stu-id="0ba03-468">1.6</span></span> |
|[<span data-ttu-id="0ba03-469">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0ba03-469">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0ba03-470">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0ba03-470">ReadItem</span></span>|
|[<span data-ttu-id="0ba03-471">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0ba03-471">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0ba03-472">読み取り</span><span class="sxs-lookup"><span data-stu-id="0ba03-472">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0ba03-473">例</span><span class="sxs-lookup"><span data-stu-id="0ba03-473">Example</span></span>

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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="0ba03-474">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="0ba03-474">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="0ba03-475">REST API または Exchange Web サービスを呼び出すために使用するトークンを含む文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="0ba03-475">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="0ba03-p132">`getCallbackTokenAsync` メソッドは、非同期の呼び出しを行なって、ユーザーのメールボックスをホストする Exchange Server から opaque トークンを取得します。コールバック トークンの有効期間は 5 分です。</span><span class="sxs-lookup"><span data-stu-id="0ba03-p132">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="0ba03-478">可能であれば、アドインでは Exchange Web サービスの代わりに REST API を使用することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="0ba03-478">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span> 

<span data-ttu-id="0ba03-479">**REST トークン**</span><span class="sxs-lookup"><span data-stu-id="0ba03-479">**REST Tokens**</span></span>

<span data-ttu-id="0ba03-p133">REST トークンが要求された場合 (`options.isRest = true`) には、作成されたトークンは Exchange Web サービスの呼び出しを認証するためには機能しません。このトークンは、アドインがマニフェストで [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) アクセス許可を指定しない限り、現在の項目およびその添付ファイルへの読み取り専用の範囲に制限されます。`ReadWriteMailbox` アクセス許可が指定された場合には、作成されるトークンは、メールを送信する機能など、メール、予定表、連絡先への読み取り/書き込みアクセスを付与します。</span><span class="sxs-lookup"><span data-stu-id="0ba03-p133">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="0ba03-483">アドインでは、`restUrl`プロパティを使用して、REST API 呼び出しを行うときに使用する正しい URL を決定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="0ba03-483">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="0ba03-484">**EWS トークン**</span><span class="sxs-lookup"><span data-stu-id="0ba03-484">**EWS Tokens**</span></span>

<span data-ttu-id="0ba03-p134">EWS トークンが要求された場合(`options.isRest = false`) には、作成されるトークンは REST API の呼び出しを認証するためには機能しません。このトークンは、現在の項目にアクセスできる範囲に制限されます。</span><span class="sxs-lookup"><span data-stu-id="0ba03-p134">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="0ba03-487">アドインでは、`ewsUrl` プロパティを使用して、EWS 呼び出しを行うときに使用する正しい URL を決定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="0ba03-487">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0ba03-488">パラメーター</span><span class="sxs-lookup"><span data-stu-id="0ba03-488">Parameters</span></span>

|<span data-ttu-id="0ba03-489">名前</span><span class="sxs-lookup"><span data-stu-id="0ba03-489">Name</span></span>| <span data-ttu-id="0ba03-490">型</span><span class="sxs-lookup"><span data-stu-id="0ba03-490">Type</span></span>| <span data-ttu-id="0ba03-491">属性</span><span class="sxs-lookup"><span data-stu-id="0ba03-491">Attributes</span></span>| <span data-ttu-id="0ba03-492">説明</span><span class="sxs-lookup"><span data-stu-id="0ba03-492">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="0ba03-493">Object</span><span class="sxs-lookup"><span data-stu-id="0ba03-493">Object</span></span> | <span data-ttu-id="0ba03-494">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="0ba03-494">&lt;optional&gt;</span></span> | <span data-ttu-id="0ba03-495">次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。</span><span class="sxs-lookup"><span data-stu-id="0ba03-495">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="0ba03-496">ブール値</span><span class="sxs-lookup"><span data-stu-id="0ba03-496">Boolean</span></span> |  <span data-ttu-id="0ba03-497">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0ba03-497">&lt;optional&gt;</span></span> | <span data-ttu-id="0ba03-p135">提供されたトークンを Outlook REST API または Exchange Web サービスに使用するかどうかを決定します。既定値は、`false`です。</span><span class="sxs-lookup"><span data-stu-id="0ba03-p135">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="0ba03-500">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="0ba03-500">Object</span></span> |  <span data-ttu-id="0ba03-501">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="0ba03-501">&lt;optional&gt;</span></span> | <span data-ttu-id="0ba03-502">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="0ba03-502">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="0ba03-503">function</span><span class="sxs-lookup"><span data-stu-id="0ba03-503">function</span></span>||<span data-ttu-id="0ba03-p136">メソッドが完了すると、`callback` パラメータに渡された関数が、単一パラメータ `asyncResult` で呼び出されます。このパラメータは、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトです。トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="0ba03-p136">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0ba03-506">要件</span><span class="sxs-lookup"><span data-stu-id="0ba03-506">Requirements</span></span>

|<span data-ttu-id="0ba03-507">要件</span><span class="sxs-lookup"><span data-stu-id="0ba03-507">Requirement</span></span>| <span data-ttu-id="0ba03-508">値</span><span class="sxs-lookup"><span data-stu-id="0ba03-508">Value</span></span>|
|---|---|
|[<span data-ttu-id="0ba03-509">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0ba03-509">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0ba03-510">1.5</span><span class="sxs-lookup"><span data-stu-id="0ba03-510">1.5</span></span> |
|[<span data-ttu-id="0ba03-511">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0ba03-511">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0ba03-512">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0ba03-512">ReadItem</span></span>|
|[<span data-ttu-id="0ba03-513">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0ba03-513">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0ba03-514">新規作成と閲覧</span><span class="sxs-lookup"><span data-stu-id="0ba03-514">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="0ba03-515">例</span><span class="sxs-lookup"><span data-stu-id="0ba03-515">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="0ba03-516">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="0ba03-516">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="0ba03-517">Exchange Server から添付ファイルやアイテムを取得するために使用するトークンを含む文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="0ba03-517">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="0ba03-p137">`getCallbackTokenAsync` メソッドは、非同期の呼び出しを行なって、ユーザーのメールボックスをホストする Exchange Server から opaque トークンを取得します。コールバック トークンの有効期間は 5 分です。</span><span class="sxs-lookup"><span data-stu-id="0ba03-p137">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="0ba03-p138">このトークンと、添付ファイル識別子または項目識別子は、サードパーティーのシステムに渡すことができます。サードパーティーのシステムでは、添付ファイルまたは項目を返すための Exchange Web サービス (EWS) の [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) または [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) 操作を呼び出すのに、トークンをベアラー承認トークンとして使用します。たとえば、リモート サービスを作成して[選択した項目から添付ファイルを取得](/outlook/add-ins/get-attachments-of-an-outlook-item)することができます。</span><span class="sxs-lookup"><span data-stu-id="0ba03-p138">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="0ba03-523">アプリでは、閲覧モードで `getCallbackTokenAsync` メソッドを呼び出すために、 **ReadItem** アクセス許可をアプリのマニフェストで指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="0ba03-523">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="0ba03-p139">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出して、`getCallbackTokenAsync` メソッドに渡すための項目識別子を取得する必要があります。アプリには、`saveAsync` メソッドを呼び出すために **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="0ba03-p139">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0ba03-526">パラメーター</span><span class="sxs-lookup"><span data-stu-id="0ba03-526">Parameters</span></span>

|<span data-ttu-id="0ba03-527">名前</span><span class="sxs-lookup"><span data-stu-id="0ba03-527">Name</span></span>| <span data-ttu-id="0ba03-528">型</span><span class="sxs-lookup"><span data-stu-id="0ba03-528">Type</span></span>| <span data-ttu-id="0ba03-529">属性</span><span class="sxs-lookup"><span data-stu-id="0ba03-529">Attributes</span></span>| <span data-ttu-id="0ba03-530">説明</span><span class="sxs-lookup"><span data-stu-id="0ba03-530">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="0ba03-531">function</span><span class="sxs-lookup"><span data-stu-id="0ba03-531">function</span></span>||<span data-ttu-id="0ba03-p140">メソッドが完了すると、`callback` パラメータに渡された関数が、単一パラメータ `asyncResult` で呼び出されます。このパラメータは、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトです。トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="0ba03-p140">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="0ba03-534">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="0ba03-534">Object</span></span>| <span data-ttu-id="0ba03-535">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="0ba03-535">&lt;optional&gt;</span></span>|<span data-ttu-id="0ba03-536">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="0ba03-536">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0ba03-537">要件</span><span class="sxs-lookup"><span data-stu-id="0ba03-537">Requirements</span></span>

|<span data-ttu-id="0ba03-538">要件</span><span class="sxs-lookup"><span data-stu-id="0ba03-538">Requirement</span></span>| <span data-ttu-id="0ba03-539">値</span><span class="sxs-lookup"><span data-stu-id="0ba03-539">Value</span></span>|
|---|---|
|[<span data-ttu-id="0ba03-540">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0ba03-540">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0ba03-541">1.3</span><span class="sxs-lookup"><span data-stu-id="0ba03-541">1.3</span></span>|
|[<span data-ttu-id="0ba03-542">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0ba03-542">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0ba03-543">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0ba03-543">ReadItem</span></span>|
|[<span data-ttu-id="0ba03-544">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0ba03-544">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0ba03-545">新規作成と閲覧</span><span class="sxs-lookup"><span data-stu-id="0ba03-545">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="0ba03-546">例</span><span class="sxs-lookup"><span data-stu-id="0ba03-546">Example</span></span>

```javascript
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="0ba03-547">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="0ba03-547">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="0ba03-548">ユーザーと Office アドインを識別するトークンを取得します。</span><span class="sxs-lookup"><span data-stu-id="0ba03-548">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="0ba03-549">`getUserIdentityTokenAsync` メソッドは、[アドインとユーザーをサードパーティのシステムで識別して認証](/outlook/add-ins/authentication)するのに使用できるトークンを返します。</span><span class="sxs-lookup"><span data-stu-id="0ba03-549">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="0ba03-550">パラメーター</span><span class="sxs-lookup"><span data-stu-id="0ba03-550">Parameters</span></span>

|<span data-ttu-id="0ba03-551">名前</span><span class="sxs-lookup"><span data-stu-id="0ba03-551">Name</span></span>| <span data-ttu-id="0ba03-552">型</span><span class="sxs-lookup"><span data-stu-id="0ba03-552">Type</span></span>| <span data-ttu-id="0ba03-553">属性</span><span class="sxs-lookup"><span data-stu-id="0ba03-553">Attributes</span></span>| <span data-ttu-id="0ba03-554">説明</span><span class="sxs-lookup"><span data-stu-id="0ba03-554">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="0ba03-555">function</span><span class="sxs-lookup"><span data-stu-id="0ba03-555">function</span></span>||<span data-ttu-id="0ba03-556">メソッドが完了すると、`callback` パラメータに渡された関数が、単一パラメータ `asyncResult` で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="0ba03-556">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="0ba03-557">トークンは、`asyncResult.value`プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="0ba03-557">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="0ba03-558">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="0ba03-558">Object</span></span>| <span data-ttu-id="0ba03-559">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="0ba03-559">&lt;optional&gt;</span></span>|<span data-ttu-id="0ba03-560">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="0ba03-560">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0ba03-561">要件</span><span class="sxs-lookup"><span data-stu-id="0ba03-561">Requirements</span></span>

|<span data-ttu-id="0ba03-562">要件</span><span class="sxs-lookup"><span data-stu-id="0ba03-562">Requirement</span></span>| <span data-ttu-id="0ba03-563">値</span><span class="sxs-lookup"><span data-stu-id="0ba03-563">Value</span></span>|
|---|---|
|[<span data-ttu-id="0ba03-564">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0ba03-564">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0ba03-565">1.0</span><span class="sxs-lookup"><span data-stu-id="0ba03-565">1.0</span></span>|
|[<span data-ttu-id="0ba03-566">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0ba03-566">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0ba03-567">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0ba03-567">ReadItem</span></span>|
|[<span data-ttu-id="0ba03-568">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0ba03-568">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0ba03-569">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0ba03-569">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0ba03-570">例</span><span class="sxs-lookup"><span data-stu-id="0ba03-570">Example</span></span>

```javascript
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="0ba03-571">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="0ba03-571">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="0ba03-572">ユーザーのメールボックスをホストしている Exchange サーバー上の Exchange Web サービス (EWS) のサービスに対して非同期の要求を行います。</span><span class="sxs-lookup"><span data-stu-id="0ba03-572">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="0ba03-573">このメソッドは、次のシナリオではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="0ba03-573">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="0ba03-574">Outlook for iOS または Outlook for Android で</span><span class="sxs-lookup"><span data-stu-id="0ba03-574">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="0ba03-575">アドインが Gmail のメールボックスにロードされる場合</span><span class="sxs-lookup"><span data-stu-id="0ba03-575">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="0ba03-576">これらの場合、アドインではユーザーのメールボックスにアクセスするために、代わりに [REST API を使用する](/outlook/add-ins/use-rest-api)必要があります。</span><span class="sxs-lookup"><span data-stu-id="0ba03-576">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="0ba03-p141">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange. See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span><span class="sxs-lookup"><span data-stu-id="0ba03-p141">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange. See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="0ba03-579">`makeEwsRequestAsync` メソッドでは、フォルダー関連アイテムを要求できません。</span><span class="sxs-lookup"><span data-stu-id="0ba03-579">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="0ba03-580">XML 要求では、UTF-8 エンコードを指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="0ba03-580">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="0ba03-p142">アドインには、`makeEwsRequestAsync` メソッドを使用するために **ReadWriteMailbox** アクセス許可が必要です。**ReadWriteMailbox** アクセス許可と、`makeEwsRequestAsync` メソッドで呼び出すことのできる EWS 操作の使用の詳細については、「[ユーザーのメールボックスへのメール アドイン アクセスのアクセス許可を指定する](/outlook/add-ins/understanding-outlook-add-in-permissions)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="0ba03-p142">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="0ba03-583">サーバー管理者は、クライアント アクセス サーバーの EWS ディレクトリで `OAuthAuthentication` を true に設定して、`makeEwsRequestAsync` メソッドで EWS 要求を行うことができるようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="0ba03-583">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="0ba03-584">バージョンの相違点</span><span class="sxs-lookup"><span data-stu-id="0ba03-584">Version differences</span></span>

<span data-ttu-id="0ba03-585">バージョン 15.0.4535.1004 より前のバージョンの Outlook で実行しているメール アプリで `makeEwsRequestAsync` メソッドを使用する場合には、エンコード値を `ISO-8859-1` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="0ba03-585">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="0ba03-p143">メール アプリが Outlook on the web で実行されている場合には、エンコード値を設定する必要はありません。メールボックスを使用してメール アプリが Outlook で実行されているのか、Outlook on the web で実行されているのかを判断する必要があります。mailbox.diagnostics.hostVersion プロパティを使用すれば、どのバージョンの Outlook が実行されているのかがわかります。</span><span class="sxs-lookup"><span data-stu-id="0ba03-p143">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0ba03-589">パラメーター</span><span class="sxs-lookup"><span data-stu-id="0ba03-589">Parameters</span></span>

|<span data-ttu-id="0ba03-590">名前</span><span class="sxs-lookup"><span data-stu-id="0ba03-590">Name</span></span>| <span data-ttu-id="0ba03-591">型</span><span class="sxs-lookup"><span data-stu-id="0ba03-591">Type</span></span>| <span data-ttu-id="0ba03-592">属性</span><span class="sxs-lookup"><span data-stu-id="0ba03-592">Attributes</span></span>| <span data-ttu-id="0ba03-593">説明</span><span class="sxs-lookup"><span data-stu-id="0ba03-593">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="0ba03-594">文字列</span><span class="sxs-lookup"><span data-stu-id="0ba03-594">String</span></span>||<span data-ttu-id="0ba03-595">EWS 要求。</span><span class="sxs-lookup"><span data-stu-id="0ba03-595">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="0ba03-596">関数</span><span class="sxs-lookup"><span data-stu-id="0ba03-596">function</span></span>||<span data-ttu-id="0ba03-597">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="0ba03-597">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="0ba03-p144">The XML result of the EWS call is provided as a string in the `asyncResult.value` property. If the result exceeds 1 MB in size, an error message is returned instead.</span><span class="sxs-lookup"><span data-stu-id="0ba03-p144">The XML result of the EWS call is provided as a string in the `asyncResult.value` property. If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="0ba03-600">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="0ba03-600">Object</span></span>| <span data-ttu-id="0ba03-601">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="0ba03-601">&lt;optional&gt;</span></span>|<span data-ttu-id="0ba03-602">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="0ba03-602">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0ba03-603">要件</span><span class="sxs-lookup"><span data-stu-id="0ba03-603">Requirements</span></span>

|<span data-ttu-id="0ba03-604">要件</span><span class="sxs-lookup"><span data-stu-id="0ba03-604">Requirement</span></span>| <span data-ttu-id="0ba03-605">値</span><span class="sxs-lookup"><span data-stu-id="0ba03-605">Value</span></span>|
|---|---|
|[<span data-ttu-id="0ba03-606">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0ba03-606">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0ba03-607">1.0</span><span class="sxs-lookup"><span data-stu-id="0ba03-607">1.0</span></span>|
|[<span data-ttu-id="0ba03-608">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0ba03-608">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0ba03-609">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="0ba03-609">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="0ba03-610">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0ba03-610">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0ba03-611">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0ba03-611">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0ba03-612">例</span><span class="sxs-lookup"><span data-stu-id="0ba03-612">Example</span></span>

<span data-ttu-id="0ba03-613">次の例は、`makeEwsRequestAsync` を呼び出し、`GetItem` 操作を使って項目の件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="0ba03-613">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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

####  <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="0ba03-614">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="0ba03-614">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="0ba03-615">サポートされているイベントの種類のイベント ハンドラーを削除します。</span><span class="sxs-lookup"><span data-stu-id="0ba03-615">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="0ba03-616">現在、サポートされているイベントの種類は `Office.EventType.ItemChanged` だけです。</span><span class="sxs-lookup"><span data-stu-id="0ba03-616">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0ba03-617">パラメーター</span><span class="sxs-lookup"><span data-stu-id="0ba03-617">Parameters</span></span>

| <span data-ttu-id="0ba03-618">名前</span><span class="sxs-lookup"><span data-stu-id="0ba03-618">Name</span></span> | <span data-ttu-id="0ba03-619">型</span><span class="sxs-lookup"><span data-stu-id="0ba03-619">Type</span></span> | <span data-ttu-id="0ba03-620">属性</span><span class="sxs-lookup"><span data-stu-id="0ba03-620">Attributes</span></span> | <span data-ttu-id="0ba03-621">説明</span><span class="sxs-lookup"><span data-stu-id="0ba03-621">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="0ba03-622">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="0ba03-622">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="0ba03-623">ハンドラーを取り消すイベント。</span><span class="sxs-lookup"><span data-stu-id="0ba03-623">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="0ba03-624">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="0ba03-624">Object</span></span> | <span data-ttu-id="0ba03-625">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="0ba03-625">&lt;optional&gt;</span></span> | <span data-ttu-id="0ba03-626">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="0ba03-626">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="0ba03-627">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="0ba03-627">Object</span></span> | <span data-ttu-id="0ba03-628">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="0ba03-628">&lt;optional&gt;</span></span> | <span data-ttu-id="0ba03-629">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="0ba03-629">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="0ba03-630">function</span><span class="sxs-lookup"><span data-stu-id="0ba03-630">function</span></span>| <span data-ttu-id="0ba03-631">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0ba03-631">&lt;optional&gt;</span></span>|<span data-ttu-id="0ba03-632">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="0ba03-632">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0ba03-633">要件</span><span class="sxs-lookup"><span data-stu-id="0ba03-633">Requirements</span></span>

|<span data-ttu-id="0ba03-634">要件</span><span class="sxs-lookup"><span data-stu-id="0ba03-634">Requirement</span></span>| <span data-ttu-id="0ba03-635">値</span><span class="sxs-lookup"><span data-stu-id="0ba03-635">Value</span></span>|
|---|---|
|[<span data-ttu-id="0ba03-636">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0ba03-636">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0ba03-637">1.5</span><span class="sxs-lookup"><span data-stu-id="0ba03-637">1.5</span></span> |
|[<span data-ttu-id="0ba03-638">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0ba03-638">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0ba03-639">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0ba03-639">ReadItem</span></span> |
|[<span data-ttu-id="0ba03-640">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0ba03-640">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0ba03-641">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0ba03-641">Compose or Read</span></span>|
