---
title: Office のメールボックス-プレビュー要件セット
description: ''
ms.date: 04/12/2019
localization_priority: Normal
ms.openlocfilehash: d19cb7c664cda42469cf7cde31d69f87101278c8
ms.sourcegitcommit: 95ed6dfbfa680dbb40ff9757020fa7e5be4760b6
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/13/2019
ms.locfileid: "31838537"
---
# <a name="mailbox"></a><span data-ttu-id="9e26b-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="9e26b-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="9e26b-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="9e26b-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="9e26b-104">Microsoft Outlook と Microsoft Outlook on the web の Outlook アドイン オブジェクト モデルへのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="9e26b-104">Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="9e26b-105">要件</span><span class="sxs-lookup"><span data-stu-id="9e26b-105">Requirements</span></span>

|<span data-ttu-id="9e26b-106">要件</span><span class="sxs-lookup"><span data-stu-id="9e26b-106">Requirement</span></span>| <span data-ttu-id="9e26b-107">値</span><span class="sxs-lookup"><span data-stu-id="9e26b-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e26b-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="9e26b-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e26b-109">1.0</span><span class="sxs-lookup"><span data-stu-id="9e26b-109">1.0</span></span>|
|[<span data-ttu-id="9e26b-110">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="9e26b-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e26b-111">制限あり</span><span class="sxs-lookup"><span data-stu-id="9e26b-111">Restricted</span></span>|
|[<span data-ttu-id="9e26b-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="9e26b-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9e26b-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="9e26b-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="9e26b-114">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="9e26b-114">Members and methods</span></span>

| <span data-ttu-id="9e26b-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="9e26b-115">Member</span></span> | <span data-ttu-id="9e26b-116">種類</span><span class="sxs-lookup"><span data-stu-id="9e26b-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="9e26b-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="9e26b-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="9e26b-118">メンバー</span><span class="sxs-lookup"><span data-stu-id="9e26b-118">Member</span></span> |
| [<span data-ttu-id="9e26b-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="9e26b-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="9e26b-120">メンバー</span><span class="sxs-lookup"><span data-stu-id="9e26b-120">Member</span></span> |
| [<span data-ttu-id="9e26b-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="9e26b-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="9e26b-122">メソッド</span><span class="sxs-lookup"><span data-stu-id="9e26b-122">Method</span></span> |
| [<span data-ttu-id="9e26b-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="9e26b-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="9e26b-124">メソッド</span><span class="sxs-lookup"><span data-stu-id="9e26b-124">Method</span></span> |
| [<span data-ttu-id="9e26b-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="9e26b-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="9e26b-126">メソッド</span><span class="sxs-lookup"><span data-stu-id="9e26b-126">Method</span></span> |
| [<span data-ttu-id="9e26b-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="9e26b-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="9e26b-128">メソッド</span><span class="sxs-lookup"><span data-stu-id="9e26b-128">Method</span></span> |
| [<span data-ttu-id="9e26b-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="9e26b-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="9e26b-130">メソッド</span><span class="sxs-lookup"><span data-stu-id="9e26b-130">Method</span></span> |
| [<span data-ttu-id="9e26b-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="9e26b-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="9e26b-132">メソッド</span><span class="sxs-lookup"><span data-stu-id="9e26b-132">Method</span></span> |
| [<span data-ttu-id="9e26b-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="9e26b-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="9e26b-134">メソッド</span><span class="sxs-lookup"><span data-stu-id="9e26b-134">Method</span></span> |
| [<span data-ttu-id="9e26b-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="9e26b-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="9e26b-136">メソッド</span><span class="sxs-lookup"><span data-stu-id="9e26b-136">Method</span></span> |
| [<span data-ttu-id="9e26b-137">displayNewMessageForm</span><span class="sxs-lookup"><span data-stu-id="9e26b-137">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="9e26b-138">メソッド</span><span class="sxs-lookup"><span data-stu-id="9e26b-138">Method</span></span> |
| [<span data-ttu-id="9e26b-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="9e26b-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="9e26b-140">メソッド</span><span class="sxs-lookup"><span data-stu-id="9e26b-140">Method</span></span> |
| [<span data-ttu-id="9e26b-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="9e26b-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="9e26b-142">メソッド</span><span class="sxs-lookup"><span data-stu-id="9e26b-142">Method</span></span> |
| [<span data-ttu-id="9e26b-143">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="9e26b-143">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="9e26b-144">メソッド</span><span class="sxs-lookup"><span data-stu-id="9e26b-144">Method</span></span> |
| [<span data-ttu-id="9e26b-145">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="9e26b-145">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="9e26b-146">メソッド</span><span class="sxs-lookup"><span data-stu-id="9e26b-146">Method</span></span> |
| [<span data-ttu-id="9e26b-147">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="9e26b-147">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="9e26b-148">メソッド</span><span class="sxs-lookup"><span data-stu-id="9e26b-148">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="9e26b-149">名前空間</span><span class="sxs-lookup"><span data-stu-id="9e26b-149">Namespaces</span></span>

<span data-ttu-id="9e26b-150">[diagnostics](Office.context.mailbox.diagnostics.md):Outlook アドインに診断情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="9e26b-150">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="9e26b-151">[item](Office.context.mailbox.item.md):Outlook アドインのメッセージや予定にアクセスするためのメソッドとプロパティを提供します。</span><span class="sxs-lookup"><span data-stu-id="9e26b-151">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="9e26b-152">[userProfile](Office.context.mailbox.userProfile.md):Outlook アドインのユーザーに関する情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="9e26b-152">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="9e26b-153">メンバー</span><span class="sxs-lookup"><span data-stu-id="9e26b-153">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="9e26b-154">ewsUrl :String</span><span class="sxs-lookup"><span data-stu-id="9e26b-154">ewsUrl :String</span></span>

<span data-ttu-id="9e26b-p101">このメール アカウントの Exchange Web サービス (EWS) エンドポイントの URL を取得します。読み取りモードのみです。</span><span class="sxs-lookup"><span data-stu-id="9e26b-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="9e26b-157">このメンバーは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="9e26b-157">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="9e26b-p102">
  `ewsUrl\` 値は、リモート サービスで、ユーザーのメールボックスに EWS 呼び出しを行うために使うことができます。たとえば、[選択したアイテムから添付ファイルを取得する](/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="9e26b-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="9e26b-160">アプリが閲覧モードで `ewsUrl` メンバーを呼び出すには、アプリのマニフェスト内に **ReadItem** アクセス許可が指定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="9e26b-160">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="9e26b-p103">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出してから、`ewsUrl` メンバーを使用する必要があります。アプリには、`saveAsync` メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="9e26b-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="9e26b-163">型</span><span class="sxs-lookup"><span data-stu-id="9e26b-163">Type</span></span>

*   <span data-ttu-id="9e26b-164">String</span><span class="sxs-lookup"><span data-stu-id="9e26b-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9e26b-165">要件</span><span class="sxs-lookup"><span data-stu-id="9e26b-165">Requirements</span></span>

|<span data-ttu-id="9e26b-166">要件</span><span class="sxs-lookup"><span data-stu-id="9e26b-166">Requirement</span></span>| <span data-ttu-id="9e26b-167">値</span><span class="sxs-lookup"><span data-stu-id="9e26b-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e26b-168">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="9e26b-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e26b-169">1.0</span><span class="sxs-lookup"><span data-stu-id="9e26b-169">1.0</span></span>|
|[<span data-ttu-id="9e26b-170">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="9e26b-170">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e26b-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9e26b-171">ReadItem</span></span>|
|[<span data-ttu-id="9e26b-172">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="9e26b-172">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9e26b-173">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="9e26b-173">Compose or Read</span></span>|

---
---

#### <a name="resturl-string"></a><span data-ttu-id="9e26b-174">restUrl :String</span><span class="sxs-lookup"><span data-stu-id="9e26b-174">restUrl :String</span></span>

<span data-ttu-id="9e26b-175">この電子メール アカウントの REST エンドポイントの URL を取得します。</span><span class="sxs-lookup"><span data-stu-id="9e26b-175">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="9e26b-176">`restUrl` 値は、ユーザーのメールボックスに [REST API](/outlook/rest/) 呼び出しを行うために使用できます。</span><span class="sxs-lookup"><span data-stu-id="9e26b-176">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="9e26b-177">アプリが閲覧モードで `restUrl` メンバーを呼び出すには、アプリのマニフェスト内に **ReadItem** アクセス許可が指定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="9e26b-177">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="9e26b-p104">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出してから、`restUrl` メンバーを使用する必要があります。アプリには、`saveAsync` メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="9e26b-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="9e26b-180">型</span><span class="sxs-lookup"><span data-stu-id="9e26b-180">Type</span></span>

*   <span data-ttu-id="9e26b-181">String</span><span class="sxs-lookup"><span data-stu-id="9e26b-181">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9e26b-182">要件</span><span class="sxs-lookup"><span data-stu-id="9e26b-182">Requirements</span></span>

|<span data-ttu-id="9e26b-183">要件</span><span class="sxs-lookup"><span data-stu-id="9e26b-183">Requirement</span></span>| <span data-ttu-id="9e26b-184">値</span><span class="sxs-lookup"><span data-stu-id="9e26b-184">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e26b-185">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="9e26b-185">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e26b-186">1.5</span><span class="sxs-lookup"><span data-stu-id="9e26b-186">1.5</span></span> |
|[<span data-ttu-id="9e26b-187">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="9e26b-187">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e26b-188">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9e26b-188">ReadItem</span></span>|
|[<span data-ttu-id="9e26b-189">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="9e26b-189">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9e26b-190">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="9e26b-190">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="9e26b-191">メソッド</span><span class="sxs-lookup"><span data-stu-id="9e26b-191">Methods</span></span>

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="9e26b-192">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="9e26b-192">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="9e26b-193">サポートされているイベントのイベント ハンドラーを追加します。</span><span class="sxs-lookup"><span data-stu-id="9e26b-193">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="9e26b-194">現在、サポートされている`Office.EventType.ItemChanged`イベント`Office.EventType.OfficeThemeChanged`の種類はおよびです。</span><span class="sxs-lookup"><span data-stu-id="9e26b-194">Currently, the supported event types are `Office.EventType.ItemChanged` and `Office.EventType.OfficeThemeChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9e26b-195">パラメーター</span><span class="sxs-lookup"><span data-stu-id="9e26b-195">Parameters</span></span>

| <span data-ttu-id="9e26b-196">名前</span><span class="sxs-lookup"><span data-stu-id="9e26b-196">Name</span></span> | <span data-ttu-id="9e26b-197">種類</span><span class="sxs-lookup"><span data-stu-id="9e26b-197">Type</span></span> | <span data-ttu-id="9e26b-198">属性</span><span class="sxs-lookup"><span data-stu-id="9e26b-198">Attributes</span></span> | <span data-ttu-id="9e26b-199">説明</span><span class="sxs-lookup"><span data-stu-id="9e26b-199">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="9e26b-200">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="9e26b-200">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="9e26b-201">ハンドラーを呼び出す必要のあるイベント。</span><span class="sxs-lookup"><span data-stu-id="9e26b-201">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="9e26b-202">関数</span><span class="sxs-lookup"><span data-stu-id="9e26b-202">Function</span></span> || <span data-ttu-id="9e26b-p105">イベントを処理する関数。関数は、オブジェクト リテラルである単一パラメーターを受け入れる必要があります。パラメーターの `type` プロパティは、`addHandlerAsync` に渡される `eventType` パラメーターと一致します。</span><span class="sxs-lookup"><span data-stu-id="9e26b-p105">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="9e26b-206">Object</span><span class="sxs-lookup"><span data-stu-id="9e26b-206">Object</span></span> | <span data-ttu-id="9e26b-207">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="9e26b-207">&lt;optional&gt;</span></span> | <span data-ttu-id="9e26b-208">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="9e26b-208">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="9e26b-209">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="9e26b-209">Object</span></span> | <span data-ttu-id="9e26b-210">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="9e26b-210">&lt;optional&gt;</span></span> | <span data-ttu-id="9e26b-211">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="9e26b-211">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="9e26b-212">function</span><span class="sxs-lookup"><span data-stu-id="9e26b-212">function</span></span>| <span data-ttu-id="9e26b-213">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="9e26b-213">&lt;optional&gt;</span></span>|<span data-ttu-id="9e26b-214">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="9e26b-214">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9e26b-215">要件</span><span class="sxs-lookup"><span data-stu-id="9e26b-215">Requirements</span></span>

|<span data-ttu-id="9e26b-216">要件</span><span class="sxs-lookup"><span data-stu-id="9e26b-216">Requirement</span></span>| <span data-ttu-id="9e26b-217">値</span><span class="sxs-lookup"><span data-stu-id="9e26b-217">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e26b-218">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="9e26b-218">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e26b-219">1.5</span><span class="sxs-lookup"><span data-stu-id="9e26b-219">1.5</span></span> |
|[<span data-ttu-id="9e26b-220">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="9e26b-220">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e26b-221">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9e26b-221">ReadItem</span></span> |
|[<span data-ttu-id="9e26b-222">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="9e26b-222">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9e26b-223">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="9e26b-223">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9e26b-224">例</span><span class="sxs-lookup"><span data-stu-id="9e26b-224">Example</span></span>

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

---
---

####  <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="9e26b-225">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="9e26b-225">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="9e26b-226">REST 形式のアイテム ID を EWS 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="9e26b-226">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="9e26b-227">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="9e26b-227">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="9e26b-p106">REST API ([Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) や [Microsoft Graph](https://graph.microsoft.io/) など) で取得されたアイテム ID は、Exchange Web サービス (EWS) に使用される形式とは異なる形式を使用します。`convertToEwsId` メソッドは、REST 形式の ID を EWS 用の適切な形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="9e26b-p106">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9e26b-230">パラメーター</span><span class="sxs-lookup"><span data-stu-id="9e26b-230">Parameters</span></span>

|<span data-ttu-id="9e26b-231">名前</span><span class="sxs-lookup"><span data-stu-id="9e26b-231">Name</span></span>| <span data-ttu-id="9e26b-232">型</span><span class="sxs-lookup"><span data-stu-id="9e26b-232">Type</span></span>| <span data-ttu-id="9e26b-233">説明</span><span class="sxs-lookup"><span data-stu-id="9e26b-233">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="9e26b-234">文字列</span><span class="sxs-lookup"><span data-stu-id="9e26b-234">String</span></span>|<span data-ttu-id="9e26b-235">Outlook REST API 形式のアイテム ID</span><span class="sxs-lookup"><span data-stu-id="9e26b-235">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="9e26b-236">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="9e26b-236">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion)|<span data-ttu-id="9e26b-237">アイテム ID の取得に使用された Outlook REST API のバージョンを示す値。</span><span class="sxs-lookup"><span data-stu-id="9e26b-237">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9e26b-238">要件</span><span class="sxs-lookup"><span data-stu-id="9e26b-238">Requirements</span></span>

|<span data-ttu-id="9e26b-239">要件</span><span class="sxs-lookup"><span data-stu-id="9e26b-239">Requirement</span></span>| <span data-ttu-id="9e26b-240">値</span><span class="sxs-lookup"><span data-stu-id="9e26b-240">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e26b-241">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="9e26b-241">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e26b-242">1.3</span><span class="sxs-lookup"><span data-stu-id="9e26b-242">1.3</span></span>|
|[<span data-ttu-id="9e26b-243">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="9e26b-243">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e26b-244">制限あり</span><span class="sxs-lookup"><span data-stu-id="9e26b-244">Restricted</span></span>|
|[<span data-ttu-id="9e26b-245">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="9e26b-245">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9e26b-246">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="9e26b-246">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9e26b-247">戻り値:</span><span class="sxs-lookup"><span data-stu-id="9e26b-247">Returns:</span></span>

<span data-ttu-id="9e26b-248">型:String</span><span class="sxs-lookup"><span data-stu-id="9e26b-248">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="9e26b-249">例</span><span class="sxs-lookup"><span data-stu-id="9e26b-249">Example</span></span>

```javascript
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

---
---

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime"></a><span data-ttu-id="9e26b-250">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="9e26b-250">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime)}</span></span>

<span data-ttu-id="9e26b-251">クライアントのローカル時間で時間情報が含まれている辞書を取得します。</span><span class="sxs-lookup"><span data-stu-id="9e26b-251">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="9e26b-p107">Outlook 用メール アプリや Outlook Web App で使う日付と時刻では、異なるタイム ゾーンを使うことができます。Outlook では、クライアント コンピューターのタイム ゾーンを使います。Outlook Web App では、Exchange 管理センター (EAC) で設定されたタイム ゾーンを使います。ユーザー インターフェイスに表示される値が、常にユーザーが期待するタイム ゾーンと一致するように日付と時刻の値を処理する必要があります。</span><span class="sxs-lookup"><span data-stu-id="9e26b-p107">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="9e26b-p108">Outlook でメール アプリが実行されている場合、`convertToLocalClientTime` メソッドは、クライアント コンピューターのタイム ゾーンに設定された値のディクショナリ オブジェクトを返します。Outlook Web Apps でメール アプリが実行されている場合、`convertToLocalClientTime` メソッドは、EAC に指定されたタイム ゾーンに設定された値のディクショナリ オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="9e26b-p108">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9e26b-257">パラメーター</span><span class="sxs-lookup"><span data-stu-id="9e26b-257">Parameters</span></span>

|<span data-ttu-id="9e26b-258">名前</span><span class="sxs-lookup"><span data-stu-id="9e26b-258">Name</span></span>| <span data-ttu-id="9e26b-259">型</span><span class="sxs-lookup"><span data-stu-id="9e26b-259">Type</span></span>| <span data-ttu-id="9e26b-260">説明</span><span class="sxs-lookup"><span data-stu-id="9e26b-260">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="9e26b-261">Date</span><span class="sxs-lookup"><span data-stu-id="9e26b-261">Date</span></span>|<span data-ttu-id="9e26b-262">日付オブジェクト</span><span class="sxs-lookup"><span data-stu-id="9e26b-262">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9e26b-263">要件</span><span class="sxs-lookup"><span data-stu-id="9e26b-263">Requirements</span></span>

|<span data-ttu-id="9e26b-264">要件</span><span class="sxs-lookup"><span data-stu-id="9e26b-264">Requirement</span></span>| <span data-ttu-id="9e26b-265">値</span><span class="sxs-lookup"><span data-stu-id="9e26b-265">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e26b-266">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="9e26b-266">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e26b-267">1.0</span><span class="sxs-lookup"><span data-stu-id="9e26b-267">1.0</span></span>|
|[<span data-ttu-id="9e26b-268">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="9e26b-268">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e26b-269">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9e26b-269">ReadItem</span></span>|
|[<span data-ttu-id="9e26b-270">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="9e26b-270">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9e26b-271">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="9e26b-271">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9e26b-272">戻り値:</span><span class="sxs-lookup"><span data-stu-id="9e26b-272">Returns:</span></span>

<span data-ttu-id="9e26b-273">型:[LocalClientTime](/javascript/api/outlook/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="9e26b-273">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime)</span></span>

---
---

####  <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="9e26b-274">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="9e26b-274">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="9e26b-275">EWS 形式のアイテム ID を REST 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="9e26b-275">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="9e26b-276">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="9e26b-276">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="9e26b-p109">EWS または `itemId` プロパティで取得されるアイテム ID は、REST API ([Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) や [Microsoft Graph](https://graph.microsoft.io/) など) に使用される形式とは異なる形式を使用します。`convertToRestId` メソッドは、EWS 形式の ID を REST 用の適切な形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="9e26b-p109">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9e26b-279">パラメーター</span><span class="sxs-lookup"><span data-stu-id="9e26b-279">Parameters</span></span>

|<span data-ttu-id="9e26b-280">名前</span><span class="sxs-lookup"><span data-stu-id="9e26b-280">Name</span></span>| <span data-ttu-id="9e26b-281">種類</span><span class="sxs-lookup"><span data-stu-id="9e26b-281">Type</span></span>| <span data-ttu-id="9e26b-282">説明</span><span class="sxs-lookup"><span data-stu-id="9e26b-282">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="9e26b-283">文字列</span><span class="sxs-lookup"><span data-stu-id="9e26b-283">String</span></span>|<span data-ttu-id="9e26b-284">Exchange Web サービス (EWS) 形式のアイテム ID</span><span class="sxs-lookup"><span data-stu-id="9e26b-284">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="9e26b-285">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="9e26b-285">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion)|<span data-ttu-id="9e26b-286">変換後の ID を使用する Outlook REST API のバージョンを示す値。</span><span class="sxs-lookup"><span data-stu-id="9e26b-286">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9e26b-287">要件</span><span class="sxs-lookup"><span data-stu-id="9e26b-287">Requirements</span></span>

|<span data-ttu-id="9e26b-288">要件</span><span class="sxs-lookup"><span data-stu-id="9e26b-288">Requirement</span></span>| <span data-ttu-id="9e26b-289">値</span><span class="sxs-lookup"><span data-stu-id="9e26b-289">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e26b-290">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="9e26b-290">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e26b-291">1.3</span><span class="sxs-lookup"><span data-stu-id="9e26b-291">1.3</span></span>|
|[<span data-ttu-id="9e26b-292">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="9e26b-292">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e26b-293">制限あり</span><span class="sxs-lookup"><span data-stu-id="9e26b-293">Restricted</span></span>|
|[<span data-ttu-id="9e26b-294">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="9e26b-294">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9e26b-295">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="9e26b-295">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9e26b-296">戻り値:</span><span class="sxs-lookup"><span data-stu-id="9e26b-296">Returns:</span></span>

<span data-ttu-id="9e26b-297">型:String</span><span class="sxs-lookup"><span data-stu-id="9e26b-297">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="9e26b-298">例</span><span class="sxs-lookup"><span data-stu-id="9e26b-298">Example</span></span>

```javascript
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

---
---

####  <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="9e26b-299">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="9e26b-299">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="9e26b-300">時間情報が含まれているディクショナリから日付オブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="9e26b-300">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="9e26b-301">`convertToUtcClientTime` メソッドは、ローカルの日付と時刻を含むディクショナリを、ローカルの日付と時刻の正しい値を持つ日付オブジェクトに変換します。</span><span class="sxs-lookup"><span data-stu-id="9e26b-301">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9e26b-302">パラメーター</span><span class="sxs-lookup"><span data-stu-id="9e26b-302">Parameters</span></span>

|<span data-ttu-id="9e26b-303">名前</span><span class="sxs-lookup"><span data-stu-id="9e26b-303">Name</span></span>| <span data-ttu-id="9e26b-304">種類</span><span class="sxs-lookup"><span data-stu-id="9e26b-304">Type</span></span>| <span data-ttu-id="9e26b-305">説明</span><span class="sxs-lookup"><span data-stu-id="9e26b-305">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="9e26b-306">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="9e26b-306">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime)|<span data-ttu-id="9e26b-307">変換するローカル時刻の値。</span><span class="sxs-lookup"><span data-stu-id="9e26b-307">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9e26b-308">要件</span><span class="sxs-lookup"><span data-stu-id="9e26b-308">Requirements</span></span>

|<span data-ttu-id="9e26b-309">要件</span><span class="sxs-lookup"><span data-stu-id="9e26b-309">Requirement</span></span>| <span data-ttu-id="9e26b-310">値</span><span class="sxs-lookup"><span data-stu-id="9e26b-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e26b-311">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="9e26b-311">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e26b-312">1.0</span><span class="sxs-lookup"><span data-stu-id="9e26b-312">1.0</span></span>|
|[<span data-ttu-id="9e26b-313">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="9e26b-313">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e26b-314">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9e26b-314">ReadItem</span></span>|
|[<span data-ttu-id="9e26b-315">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="9e26b-315">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9e26b-316">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="9e26b-316">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9e26b-317">戻り値:</span><span class="sxs-lookup"><span data-stu-id="9e26b-317">Returns:</span></span>

<span data-ttu-id="9e26b-318">時間が UTC で表現された日付オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="9e26b-318">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="9e26b-319">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="9e26b-319">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="9e26b-320">日付</span><span class="sxs-lookup"><span data-stu-id="9e26b-320">Date</span></span></dd>

</dl>

---
---

####  <a name="displayappointmentformitemid"></a><span data-ttu-id="9e26b-321">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="9e26b-321">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="9e26b-322">既存の予定を表示します。</span><span class="sxs-lookup"><span data-stu-id="9e26b-322">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="9e26b-323">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="9e26b-323">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="9e26b-324">`displayAppointmentForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存の予定を開きます。</span><span class="sxs-lookup"><span data-stu-id="9e26b-324">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="9e26b-p110">Outlook for Mac では、この方法を使って、定期的な系列の一部ではない単一の予定や定期的な系列のマスター予定を表示できます。ただし、系列のインスタンスは表示できません。これは、Outlook for Mac においては定期的な系列のインスタンスのプロパティ (アイテム ID を含む) にアクセスできないためです。</span><span class="sxs-lookup"><span data-stu-id="9e26b-p110">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="9e26b-327">Outlook Web App では、このメソッドは指定されたフォームの本文が 32 KB 以下の文字数の場合にフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="9e26b-327">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="9e26b-328">指定のアイテム識別子が既存の予定を表していない場合は、クライアント コンピューターまたはデバイスで空のウィンドウが開き、エラー メッセージは返されません。</span><span class="sxs-lookup"><span data-stu-id="9e26b-328">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9e26b-329">パラメーター</span><span class="sxs-lookup"><span data-stu-id="9e26b-329">Parameters</span></span>

|<span data-ttu-id="9e26b-330">名前</span><span class="sxs-lookup"><span data-stu-id="9e26b-330">Name</span></span>| <span data-ttu-id="9e26b-331">種類</span><span class="sxs-lookup"><span data-stu-id="9e26b-331">Type</span></span>| <span data-ttu-id="9e26b-332">説明</span><span class="sxs-lookup"><span data-stu-id="9e26b-332">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="9e26b-333">文字列</span><span class="sxs-lookup"><span data-stu-id="9e26b-333">String</span></span>|<span data-ttu-id="9e26b-334">既存の予定の Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="9e26b-334">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9e26b-335">要件</span><span class="sxs-lookup"><span data-stu-id="9e26b-335">Requirements</span></span>

|<span data-ttu-id="9e26b-336">要件</span><span class="sxs-lookup"><span data-stu-id="9e26b-336">Requirement</span></span>| <span data-ttu-id="9e26b-337">値</span><span class="sxs-lookup"><span data-stu-id="9e26b-337">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e26b-338">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="9e26b-338">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e26b-339">1.0</span><span class="sxs-lookup"><span data-stu-id="9e26b-339">1.0</span></span>|
|[<span data-ttu-id="9e26b-340">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="9e26b-340">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e26b-341">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9e26b-341">ReadItem</span></span>|
|[<span data-ttu-id="9e26b-342">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="9e26b-342">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9e26b-343">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="9e26b-343">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9e26b-344">例</span><span class="sxs-lookup"><span data-stu-id="9e26b-344">Example</span></span>

```javascript
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

---
---

####  <a name="displaymessageformitemid"></a><span data-ttu-id="9e26b-345">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="9e26b-345">displayMessageForm(itemId)</span></span>

<span data-ttu-id="9e26b-346">既存のメッセージを表示します。</span><span class="sxs-lookup"><span data-stu-id="9e26b-346">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="9e26b-347">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="9e26b-347">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="9e26b-348">`displayMessageForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存のメッセージを開きます。</span><span class="sxs-lookup"><span data-stu-id="9e26b-348">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="9e26b-349">Outlook Web App では、このメソッドは指定されたフォームの本文が 32 KB 以下の文字数の場合にフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="9e26b-349">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="9e26b-350">指定のアイテム識別子が既存のメッセージを表していない場合は、クライアント コンピューターにはメッセージは表示されず、エラー メッセージも返されません。</span><span class="sxs-lookup"><span data-stu-id="9e26b-350">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="9e26b-p111">予定を表す `itemId` を含む `displayMessageForm` を使用しないでください。既存の予定を表示するには、`displayAppointmentForm` メソッドを使用します。新しい予定を作成するフォームを表示するには、`displayNewAppointmentForm` メソッドを使用します。</span><span class="sxs-lookup"><span data-stu-id="9e26b-p111">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9e26b-353">パラメーター</span><span class="sxs-lookup"><span data-stu-id="9e26b-353">Parameters</span></span>

|<span data-ttu-id="9e26b-354">名前</span><span class="sxs-lookup"><span data-stu-id="9e26b-354">Name</span></span>| <span data-ttu-id="9e26b-355">種類</span><span class="sxs-lookup"><span data-stu-id="9e26b-355">Type</span></span>| <span data-ttu-id="9e26b-356">説明</span><span class="sxs-lookup"><span data-stu-id="9e26b-356">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="9e26b-357">文字列</span><span class="sxs-lookup"><span data-stu-id="9e26b-357">String</span></span>|<span data-ttu-id="9e26b-358">既存のメッセージの Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="9e26b-358">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9e26b-359">要件</span><span class="sxs-lookup"><span data-stu-id="9e26b-359">Requirements</span></span>

|<span data-ttu-id="9e26b-360">要件</span><span class="sxs-lookup"><span data-stu-id="9e26b-360">Requirement</span></span>| <span data-ttu-id="9e26b-361">値</span><span class="sxs-lookup"><span data-stu-id="9e26b-361">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e26b-362">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="9e26b-362">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e26b-363">1.0</span><span class="sxs-lookup"><span data-stu-id="9e26b-363">1.0</span></span>|
|[<span data-ttu-id="9e26b-364">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="9e26b-364">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e26b-365">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9e26b-365">ReadItem</span></span>|
|[<span data-ttu-id="9e26b-366">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="9e26b-366">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9e26b-367">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="9e26b-367">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9e26b-368">例</span><span class="sxs-lookup"><span data-stu-id="9e26b-368">Example</span></span>

```javascript
Office.context.mailbox.displayMessageForm(messageId);
```

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="9e26b-369">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="9e26b-369">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="9e26b-370">新しい予定を作成するためのフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="9e26b-370">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="9e26b-371">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="9e26b-371">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="9e26b-p112">`displayNewAppointmentForm` メソッドを使用すると、ユーザーが新しい予定または会議を作成できるフォームが開きます。パラメーターを指定すると、予定のフォーム フィールドにパラメーターの内容が自動的に設定されます。</span><span class="sxs-lookup"><span data-stu-id="9e26b-p112">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="9e26b-p113">このメソッドは、Outlook Web App と OWA for Devices において、出席者フィールドが含まれるフォームを必ず表示します。入力引数として出席者を指定しないと、このメソッドにより **[保存]** ボタンのあるフォームが表示されます。出席者を指定した場合には、フォームにその出席者と **[送信]** ボタンが表示されます。</span><span class="sxs-lookup"><span data-stu-id="9e26b-p113">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="9e26b-p114">Outlook リッチ クライアントと Outlook RT で、`requiredAttendees`、`optionalAttendees`、または `resources` パラメーターに出席者またはリソースを指定し、このメソッドを実行すると、**[送信]** ボタンがある会議フォームが表示されます。受信者を指定せずにこのメソッドを実行すると、**[保存して閉じる]** ボタンがある予定フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="9e26b-p114">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="9e26b-379">パラメーターのいずれかが指定のサイズ制限を超える場合、または不明なパラメーター名が指定されている場合は、例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="9e26b-379">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9e26b-380">パラメーター</span><span class="sxs-lookup"><span data-stu-id="9e26b-380">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="9e26b-381">すべてのパラメーターは省略可能です。</span><span class="sxs-lookup"><span data-stu-id="9e26b-381">All parameters are optional.</span></span>

|<span data-ttu-id="9e26b-382">名前</span><span class="sxs-lookup"><span data-stu-id="9e26b-382">Name</span></span>| <span data-ttu-id="9e26b-383">型</span><span class="sxs-lookup"><span data-stu-id="9e26b-383">Type</span></span>| <span data-ttu-id="9e26b-384">説明</span><span class="sxs-lookup"><span data-stu-id="9e26b-384">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="9e26b-385">Object</span><span class="sxs-lookup"><span data-stu-id="9e26b-385">Object</span></span> | <span data-ttu-id="9e26b-386">新しい予定を記述するパラメーターのディクショナリ。</span><span class="sxs-lookup"><span data-stu-id="9e26b-386">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="9e26b-387">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="9e26b-387">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="9e26b-p115">予定に必要な各出席者について、メール アドレスを含む文字列の配列、または `EmailAddressDetails` オブジェクトを含む配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="9e26b-p115">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="9e26b-390">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="9e26b-390">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="9e26b-p116">予定の各任意出席者について、メール アドレスを含む文字列の配列、または `EmailAddressDetails` オブジェクトを含む配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="9e26b-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="9e26b-393">日付</span><span class="sxs-lookup"><span data-stu-id="9e26b-393">Date</span></span> | <span data-ttu-id="9e26b-394">予定の開始日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="9e26b-394">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="9e26b-395">日付</span><span class="sxs-lookup"><span data-stu-id="9e26b-395">Date</span></span> | <span data-ttu-id="9e26b-396">予定の終了日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="9e26b-396">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="9e26b-397">String</span><span class="sxs-lookup"><span data-stu-id="9e26b-397">String</span></span> | <span data-ttu-id="9e26b-p117">予定の場所を含む文字列。文字列は最大 255 文字に制限されます。</span><span class="sxs-lookup"><span data-stu-id="9e26b-p117">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="9e26b-400">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="9e26b-400">Array.&lt;String&gt;</span></span> | <span data-ttu-id="9e26b-p118">予定に必要なリソースを含む文字列の配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="9e26b-p118">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="9e26b-403">文字列</span><span class="sxs-lookup"><span data-stu-id="9e26b-403">String</span></span> | <span data-ttu-id="9e26b-p119">予定の件名を含む文字列です。文字列は最大 255 文字に制限されます。</span><span class="sxs-lookup"><span data-stu-id="9e26b-p119">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="9e26b-406">String</span><span class="sxs-lookup"><span data-stu-id="9e26b-406">String</span></span> | <span data-ttu-id="9e26b-p120">予定の本文。本文の内容は、最大サイズが 32 KB に制限されます。</span><span class="sxs-lookup"><span data-stu-id="9e26b-p120">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="9e26b-409">要件</span><span class="sxs-lookup"><span data-stu-id="9e26b-409">Requirements</span></span>

|<span data-ttu-id="9e26b-410">要件</span><span class="sxs-lookup"><span data-stu-id="9e26b-410">Requirement</span></span>| <span data-ttu-id="9e26b-411">値</span><span class="sxs-lookup"><span data-stu-id="9e26b-411">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e26b-412">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="9e26b-412">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e26b-413">1.0</span><span class="sxs-lookup"><span data-stu-id="9e26b-413">1.0</span></span>|
|[<span data-ttu-id="9e26b-414">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="9e26b-414">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e26b-415">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9e26b-415">ReadItem</span></span>|
|[<span data-ttu-id="9e26b-416">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="9e26b-416">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9e26b-417">読み取り</span><span class="sxs-lookup"><span data-stu-id="9e26b-417">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9e26b-418">例</span><span class="sxs-lookup"><span data-stu-id="9e26b-418">Example</span></span>

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

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="9e26b-419">displayNewMessageForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="9e26b-419">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="9e26b-420">新しいメッセージを作成するためのフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="9e26b-420">Displays a form for creating a new message.</span></span>

<span data-ttu-id="9e26b-p121">The `displayNewMessageForm` method opens a form that enables the user to create a new message. If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span><span class="sxs-lookup"><span data-stu-id="9e26b-p121">The `displayNewMessageForm` method opens a form that enables the user to create a new message. If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="9e26b-423">パラメーターのいずれかが指定のサイズ制限を超える場合、または不明なパラメーター名が指定されている場合は、例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="9e26b-423">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9e26b-424">パラメーター</span><span class="sxs-lookup"><span data-stu-id="9e26b-424">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="9e26b-425">すべてのパラメーターは省略可能です。</span><span class="sxs-lookup"><span data-stu-id="9e26b-425">All parameters are optional.</span></span>

|<span data-ttu-id="9e26b-426">名前</span><span class="sxs-lookup"><span data-stu-id="9e26b-426">Name</span></span>| <span data-ttu-id="9e26b-427">型</span><span class="sxs-lookup"><span data-stu-id="9e26b-427">Type</span></span>| <span data-ttu-id="9e26b-428">説明</span><span class="sxs-lookup"><span data-stu-id="9e26b-428">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="9e26b-429">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="9e26b-429">Object</span></span> | <span data-ttu-id="9e26b-430">新しいメッセージを記述するパラメータの辞書。</span><span class="sxs-lookup"><span data-stu-id="9e26b-430">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="9e26b-431">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="9e26b-431">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="9e26b-p122">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line. The array is limited to a maximum of 100 entries.</span><span class="sxs-lookup"><span data-stu-id="9e26b-p122">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="9e26b-434">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="9e26b-434">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="9e26b-p123">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line. The array is limited to a maximum of 100 entries.</span><span class="sxs-lookup"><span data-stu-id="9e26b-p123">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="9e26b-437">配列。&lt;文字列&gt; | 配列。&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="9e26b-437">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="9e26b-p124">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line. The array is limited to a maximum of 100 entries.</span><span class="sxs-lookup"><span data-stu-id="9e26b-p124">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="9e26b-440">文字列</span><span class="sxs-lookup"><span data-stu-id="9e26b-440">String</span></span> | <span data-ttu-id="9e26b-p125">A string containing the subject of the message. The string is limited to a maximum of 255 characters.</span><span class="sxs-lookup"><span data-stu-id="9e26b-p125">A string containing the subject of the message. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="9e26b-443">文字列</span><span class="sxs-lookup"><span data-stu-id="9e26b-443">String</span></span> | <span data-ttu-id="9e26b-p126">The HTML body of the message. The body content is limited to a maximum size of 32 KB.</span><span class="sxs-lookup"><span data-stu-id="9e26b-p126">The HTML body of the message. The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="9e26b-446">配列。&lt;オブジェクト&gt;</span><span class="sxs-lookup"><span data-stu-id="9e26b-446">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="9e26b-447">添付ファイルまたは添付アイテムである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="9e26b-447">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="9e26b-448">文字列</span><span class="sxs-lookup"><span data-stu-id="9e26b-448">String</span></span> | <span data-ttu-id="9e26b-p127">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="9e26b-p127">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="9e26b-451">String</span><span class="sxs-lookup"><span data-stu-id="9e26b-451">String</span></span> | <span data-ttu-id="9e26b-452">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="9e26b-452">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="9e26b-453">文字列</span><span class="sxs-lookup"><span data-stu-id="9e26b-453">String</span></span> | <span data-ttu-id="9e26b-p128">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="9e26b-p128">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="9e26b-456">ブール値</span><span class="sxs-lookup"><span data-stu-id="9e26b-456">Boolean</span></span> | <span data-ttu-id="9e26b-p129">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="9e26b-p129">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="9e26b-459">文字列</span><span class="sxs-lookup"><span data-stu-id="9e26b-459">String</span></span> | <span data-ttu-id="9e26b-p130">Only used if `type` is set to `item`. The EWS item id of the existing e-mail you want to attach to the new message. This is a string up to 100 characters.</span><span class="sxs-lookup"><span data-stu-id="9e26b-p130">Only used if `type` is set to `item`. The EWS item id of the existing e-mail you want to attach to the new message. This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="9e26b-463">要件</span><span class="sxs-lookup"><span data-stu-id="9e26b-463">Requirements</span></span>

|<span data-ttu-id="9e26b-464">要件</span><span class="sxs-lookup"><span data-stu-id="9e26b-464">Requirement</span></span>| <span data-ttu-id="9e26b-465">値</span><span class="sxs-lookup"><span data-stu-id="9e26b-465">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e26b-466">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="9e26b-466">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e26b-467">1.6</span><span class="sxs-lookup"><span data-stu-id="9e26b-467">1.6</span></span> |
|[<span data-ttu-id="9e26b-468">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="9e26b-468">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e26b-469">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9e26b-469">ReadItem</span></span>|
|[<span data-ttu-id="9e26b-470">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="9e26b-470">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9e26b-471">読み取り</span><span class="sxs-lookup"><span data-stu-id="9e26b-471">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9e26b-472">例</span><span class="sxs-lookup"><span data-stu-id="9e26b-472">Example</span></span>

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

---
---

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="9e26b-473">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="9e26b-473">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="9e26b-474">REST API または Exchange Web サービスを呼び出すために使用するトークンを含む文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="9e26b-474">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="9e26b-p131">`getCallbackTokenAsync` メソッドは、ユーザーのメールボックスをホストする Exchange Server から不透明なトークンを取得する非同期の呼び出しを行います。コールバック トークンの有効期間は 5 分です。</span><span class="sxs-lookup"><span data-stu-id="9e26b-p131">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="9e26b-477">可能であれば、アドインでは Exchange Web サービスの代わりに REST API を使用することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="9e26b-477">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span>

<span data-ttu-id="9e26b-478">**REST トークン**</span><span class="sxs-lookup"><span data-stu-id="9e26b-478">**REST Tokens**</span></span>

<span data-ttu-id="9e26b-p132">REST トークンが要求された場合 (`options.isRest = true`)、結果トークンは Exchange Web サービスの呼び出しを認証するためには機能しません。アドインがマニフェストで [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) アクセス許可を指定していない限り、トークンの範囲は現在のアイテムとその添付ファイルへの読み取り専用アクセスに制限されます。`ReadWriteMailbox` アクセス許可が指定されている場合は、結果トークンは、メールを送信する機能など、メール、カレンダー、連絡先への読み取り/書き込みアクセスを付与します。</span><span class="sxs-lookup"><span data-stu-id="9e26b-p132">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="9e26b-482">アドインでは、`restUrl` プロパティを使用して、REST API 呼び出しを行うときに使用する正しい URL を決定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="9e26b-482">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="9e26b-483">**EWS トークン**</span><span class="sxs-lookup"><span data-stu-id="9e26b-483">**EWS Tokens**</span></span>

<span data-ttu-id="9e26b-p133">EWS トークンが要求された場合 (`options.isRest = false`)、結果トークンは REST API 呼び出しを認証するためには機能しません。トークンの範囲は、現在のアイテムへのアクセスに制限されます。</span><span class="sxs-lookup"><span data-stu-id="9e26b-p133">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="9e26b-486">アドインでは、`ewsUrl` プロパティを使用して、EWS 呼び出しを行うときに使用する正しい URL を決定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="9e26b-486">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9e26b-487">パラメーター</span><span class="sxs-lookup"><span data-stu-id="9e26b-487">Parameters</span></span>

|<span data-ttu-id="9e26b-488">名前</span><span class="sxs-lookup"><span data-stu-id="9e26b-488">Name</span></span>| <span data-ttu-id="9e26b-489">型</span><span class="sxs-lookup"><span data-stu-id="9e26b-489">Type</span></span>| <span data-ttu-id="9e26b-490">属性</span><span class="sxs-lookup"><span data-stu-id="9e26b-490">Attributes</span></span>| <span data-ttu-id="9e26b-491">説明</span><span class="sxs-lookup"><span data-stu-id="9e26b-491">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="9e26b-492">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="9e26b-492">Object</span></span> | <span data-ttu-id="9e26b-493">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="9e26b-493">&lt;optional&gt;</span></span> | <span data-ttu-id="9e26b-494">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="9e26b-494">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="9e26b-495">Boolean</span><span class="sxs-lookup"><span data-stu-id="9e26b-495">Boolean</span></span> |  <span data-ttu-id="9e26b-496">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="9e26b-496">&lt;optional&gt;</span></span> | <span data-ttu-id="9e26b-p134">提供されたトークンを Outlook REST API または Exchange Web サービスに使用するかどうかを決定します。既定値は、`false` です。</span><span class="sxs-lookup"><span data-stu-id="9e26b-p134">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="9e26b-499">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="9e26b-499">Object</span></span> |  <span data-ttu-id="9e26b-500">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="9e26b-500">&lt;optional&gt;</span></span> | <span data-ttu-id="9e26b-501">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="9e26b-501">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="9e26b-502">function</span><span class="sxs-lookup"><span data-stu-id="9e26b-502">function</span></span>||<span data-ttu-id="9e26b-p135">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="9e26b-p135">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9e26b-505">要件</span><span class="sxs-lookup"><span data-stu-id="9e26b-505">Requirements</span></span>

|<span data-ttu-id="9e26b-506">要件</span><span class="sxs-lookup"><span data-stu-id="9e26b-506">Requirement</span></span>| <span data-ttu-id="9e26b-507">値</span><span class="sxs-lookup"><span data-stu-id="9e26b-507">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e26b-508">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="9e26b-508">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e26b-509">1.5</span><span class="sxs-lookup"><span data-stu-id="9e26b-509">1.5</span></span> |
|[<span data-ttu-id="9e26b-510">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="9e26b-510">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e26b-511">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9e26b-511">ReadItem</span></span>|
|[<span data-ttu-id="9e26b-512">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="9e26b-512">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9e26b-513">新規作成と閲覧</span><span class="sxs-lookup"><span data-stu-id="9e26b-513">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="9e26b-514">例</span><span class="sxs-lookup"><span data-stu-id="9e26b-514">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="9e26b-515">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="9e26b-515">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="9e26b-516">Exchange Server から添付ファイルやアイテムを取得するために使うトークンを含む文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="9e26b-516">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="9e26b-p136">`getCallbackTokenAsync` メソッドは、ユーザーのメールボックスをホストする Exchange Server から不透明なトークンを取得する非同期の呼び出しを行います。コールバック トークンの有効期間は 5 分です。</span><span class="sxs-lookup"><span data-stu-id="9e26b-p136">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="9e26b-p137">トークンと予定の識別子またはアイテムの識別子をサードパーティ システムに渡すことができます。サードパーティ システムは、トークンをベアラー承認トークンとして使用し、Exchange Web サービス (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) または [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) 操作を呼び出して、添付ファイルまたはアイテムを返します。たとえば、[選択したアイテムから添付ファイルを取得する](/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="9e26b-p137">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="9e26b-522">アプリが閲覧モードで `getCallbackTokenAsync` メソッドを呼び出すには、アプリのマニフェスト内に **ReadItem** アクセス許可が指定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="9e26b-522">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="9e26b-p138">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出してアイテムの識別子を `getCallbackTokenAsync` メソッドに渡す必要があります。アプリには、`saveAsync` メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="9e26b-p138">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9e26b-525">パラメーター</span><span class="sxs-lookup"><span data-stu-id="9e26b-525">Parameters</span></span>

|<span data-ttu-id="9e26b-526">名前</span><span class="sxs-lookup"><span data-stu-id="9e26b-526">Name</span></span>| <span data-ttu-id="9e26b-527">型</span><span class="sxs-lookup"><span data-stu-id="9e26b-527">Type</span></span>| <span data-ttu-id="9e26b-528">属性</span><span class="sxs-lookup"><span data-stu-id="9e26b-528">Attributes</span></span>| <span data-ttu-id="9e26b-529">説明</span><span class="sxs-lookup"><span data-stu-id="9e26b-529">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="9e26b-530">function</span><span class="sxs-lookup"><span data-stu-id="9e26b-530">function</span></span>||<span data-ttu-id="9e26b-p139">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="9e26b-p139">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="9e26b-533">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="9e26b-533">Object</span></span>| <span data-ttu-id="9e26b-534">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="9e26b-534">&lt;optional&gt;</span></span>|<span data-ttu-id="9e26b-535">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="9e26b-535">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9e26b-536">要件</span><span class="sxs-lookup"><span data-stu-id="9e26b-536">Requirements</span></span>

|<span data-ttu-id="9e26b-537">要件</span><span class="sxs-lookup"><span data-stu-id="9e26b-537">Requirement</span></span>| <span data-ttu-id="9e26b-538">値</span><span class="sxs-lookup"><span data-stu-id="9e26b-538">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e26b-539">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="9e26b-539">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e26b-540">1.3</span><span class="sxs-lookup"><span data-stu-id="9e26b-540">1.3</span></span>|
|[<span data-ttu-id="9e26b-541">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="9e26b-541">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e26b-542">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9e26b-542">ReadItem</span></span>|
|[<span data-ttu-id="9e26b-543">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="9e26b-543">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9e26b-544">新規作成と閲覧</span><span class="sxs-lookup"><span data-stu-id="9e26b-544">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="9e26b-545">例</span><span class="sxs-lookup"><span data-stu-id="9e26b-545">Example</span></span>

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

####  <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="9e26b-546">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="9e26b-546">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="9e26b-547">ユーザーと Office アドインを識別するトークンを取得します。</span><span class="sxs-lookup"><span data-stu-id="9e26b-547">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="9e26b-548">`getUserIdentityTokenAsync` メソッドは、[アドインとユーザーをサード パーティのシステムで識別して認証](/outlook/add-ins/authentication)することのできるトークンを返します。</span><span class="sxs-lookup"><span data-stu-id="9e26b-548">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="9e26b-549">パラメーター</span><span class="sxs-lookup"><span data-stu-id="9e26b-549">Parameters</span></span>

|<span data-ttu-id="9e26b-550">名前</span><span class="sxs-lookup"><span data-stu-id="9e26b-550">Name</span></span>| <span data-ttu-id="9e26b-551">型</span><span class="sxs-lookup"><span data-stu-id="9e26b-551">Type</span></span>| <span data-ttu-id="9e26b-552">属性</span><span class="sxs-lookup"><span data-stu-id="9e26b-552">Attributes</span></span>| <span data-ttu-id="9e26b-553">説明</span><span class="sxs-lookup"><span data-stu-id="9e26b-553">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="9e26b-554">関数</span><span class="sxs-lookup"><span data-stu-id="9e26b-554">function</span></span>||<span data-ttu-id="9e26b-555">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="9e26b-555">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="9e26b-556">トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="9e26b-556">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="9e26b-557">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="9e26b-557">Object</span></span>| <span data-ttu-id="9e26b-558">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="9e26b-558">&lt;optional&gt;</span></span>|<span data-ttu-id="9e26b-559">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="9e26b-559">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9e26b-560">要件</span><span class="sxs-lookup"><span data-stu-id="9e26b-560">Requirements</span></span>

|<span data-ttu-id="9e26b-561">要件</span><span class="sxs-lookup"><span data-stu-id="9e26b-561">Requirement</span></span>| <span data-ttu-id="9e26b-562">値</span><span class="sxs-lookup"><span data-stu-id="9e26b-562">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e26b-563">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="9e26b-563">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e26b-564">1.0</span><span class="sxs-lookup"><span data-stu-id="9e26b-564">1.0</span></span>|
|[<span data-ttu-id="9e26b-565">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="9e26b-565">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e26b-566">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9e26b-566">ReadItem</span></span>|
|[<span data-ttu-id="9e26b-567">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="9e26b-567">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9e26b-568">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="9e26b-568">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9e26b-569">例</span><span class="sxs-lookup"><span data-stu-id="9e26b-569">Example</span></span>

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

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="9e26b-570">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="9e26b-570">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="9e26b-571">ユーザーのメールボックスをホストしている Exchange サーバー上の Exchange Web サービス (EWS) のサービスに対して非同期の要求を行います。</span><span class="sxs-lookup"><span data-stu-id="9e26b-571">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="9e26b-572">このメソッドは、次のシナリオではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="9e26b-572">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="9e26b-573">Outlook for iOS または Outlook for Android を使用している場合</span><span class="sxs-lookup"><span data-stu-id="9e26b-573">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="9e26b-574">アドインが Gmail のメールボックスに読み込まれる場合</span><span class="sxs-lookup"><span data-stu-id="9e26b-574">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="9e26b-575">このような場合は、アドインでは [REST API を使用](/outlook/add-ins/use-rest-api)して、代わりにユーザーのメールボックスにアクセスする必要があります。</span><span class="sxs-lookup"><span data-stu-id="9e26b-575">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="9e26b-p140">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange. See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span><span class="sxs-lookup"><span data-stu-id="9e26b-p140">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange. See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="9e26b-578">`makeEwsRequestAsync` メソッドでは、フォルダー関連アイテムを要求できません。</span><span class="sxs-lookup"><span data-stu-id="9e26b-578">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="9e26b-579">XML 要求では UTF-8 エンコードを指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="9e26b-579">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="9e26b-p141">`makeEwsRequestAsync` メソッドを使用するには、アドインに **ReadWriteMailbox** アクセス許可が必要です。**ReadWriteMailbox** アクセス許可と、`makeEwsRequestAsync` メソッドで呼び出せる EWS 操作の使い方については、「[ユーザーのメールボックスへのメール アドイン アクセスのアクセス許可を指定する](/outlook/add-ins/understanding-outlook-add-in-permissions)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="9e26b-p141">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="9e26b-582">サーバー管理者は、クライアント アクセス サーバーの EWS ディレクトリで `OAuthAuthentication` を true に設定して、`makeEwsRequestAsync` メソッドで EWS 要求を行うことができるようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="9e26b-582">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="9e26b-583">バージョンの相違点</span><span class="sxs-lookup"><span data-stu-id="9e26b-583">Version differences</span></span>

<span data-ttu-id="9e26b-584">バージョン 15.0.4535.1004 より前のバージョンの Outlook で実行しているメール アプリで `makeEwsRequestAsync` メソッドを使う場合は、エンコード値を `ISO-8859-1` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="9e26b-584">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="9e26b-p142">Outlook on the web でメール アプリを実行している場合は、エンコード値を設定する必要はありません。mailbox.diagnostics.hostName プロパティを使って、メール アプリを Outlook で実行しているのか、Outlook on the web で実行しているのかを確認できます。mailbox.diagnostics.hostVersion プロパティを使って、どのバージョンの Outlook を使って実行しているかを確認できます。</span><span class="sxs-lookup"><span data-stu-id="9e26b-p142">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9e26b-588">パラメーター</span><span class="sxs-lookup"><span data-stu-id="9e26b-588">Parameters</span></span>

|<span data-ttu-id="9e26b-589">名前</span><span class="sxs-lookup"><span data-stu-id="9e26b-589">Name</span></span>| <span data-ttu-id="9e26b-590">型</span><span class="sxs-lookup"><span data-stu-id="9e26b-590">Type</span></span>| <span data-ttu-id="9e26b-591">属性</span><span class="sxs-lookup"><span data-stu-id="9e26b-591">Attributes</span></span>| <span data-ttu-id="9e26b-592">説明</span><span class="sxs-lookup"><span data-stu-id="9e26b-592">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="9e26b-593">文字列</span><span class="sxs-lookup"><span data-stu-id="9e26b-593">String</span></span>||<span data-ttu-id="9e26b-594">EWS 要求です。</span><span class="sxs-lookup"><span data-stu-id="9e26b-594">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="9e26b-595">関数</span><span class="sxs-lookup"><span data-stu-id="9e26b-595">function</span></span>||<span data-ttu-id="9e26b-596">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="9e26b-596">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="9e26b-p143">The XML result of the EWS call is provided as a string in the `asyncResult.value` property. If the result exceeds 1 MB in size, an error message is returned instead.</span><span class="sxs-lookup"><span data-stu-id="9e26b-p143">The XML result of the EWS call is provided as a string in the `asyncResult.value` property. If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="9e26b-599">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="9e26b-599">Object</span></span>| <span data-ttu-id="9e26b-600">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="9e26b-600">&lt;optional&gt;</span></span>|<span data-ttu-id="9e26b-601">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="9e26b-601">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9e26b-602">要件</span><span class="sxs-lookup"><span data-stu-id="9e26b-602">Requirements</span></span>

|<span data-ttu-id="9e26b-603">要件</span><span class="sxs-lookup"><span data-stu-id="9e26b-603">Requirement</span></span>| <span data-ttu-id="9e26b-604">値</span><span class="sxs-lookup"><span data-stu-id="9e26b-604">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e26b-605">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="9e26b-605">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e26b-606">1.0</span><span class="sxs-lookup"><span data-stu-id="9e26b-606">1.0</span></span>|
|[<span data-ttu-id="9e26b-607">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="9e26b-607">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e26b-608">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="9e26b-608">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="9e26b-609">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="9e26b-609">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9e26b-610">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="9e26b-610">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9e26b-611">例</span><span class="sxs-lookup"><span data-stu-id="9e26b-611">Example</span></span>

<span data-ttu-id="9e26b-612">次の例は、`GetItem` 操作を使ってアイテムの件名を取得するため、`makeEwsRequestAsync` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="9e26b-612">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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

####  <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="9e26b-613">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="9e26b-613">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="9e26b-614">サポートされているイベントの種類のイベント ハンドラーを削除します。</span><span class="sxs-lookup"><span data-stu-id="9e26b-614">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="9e26b-615">現在、サポートされている`Office.EventType.ItemChanged`イベント`Office.EventType.OfficeThemeChanged`の種類はおよびです。</span><span class="sxs-lookup"><span data-stu-id="9e26b-615">Currently, the supported event types are `Office.EventType.ItemChanged` and `Office.EventType.OfficeThemeChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9e26b-616">パラメーター</span><span class="sxs-lookup"><span data-stu-id="9e26b-616">Parameters</span></span>

| <span data-ttu-id="9e26b-617">名前</span><span class="sxs-lookup"><span data-stu-id="9e26b-617">Name</span></span> | <span data-ttu-id="9e26b-618">型</span><span class="sxs-lookup"><span data-stu-id="9e26b-618">Type</span></span> | <span data-ttu-id="9e26b-619">属性</span><span class="sxs-lookup"><span data-stu-id="9e26b-619">Attributes</span></span> | <span data-ttu-id="9e26b-620">説明</span><span class="sxs-lookup"><span data-stu-id="9e26b-620">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="9e26b-621">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="9e26b-621">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="9e26b-622">ハンドラーを取り消すイベント。</span><span class="sxs-lookup"><span data-stu-id="9e26b-622">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="9e26b-623">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="9e26b-623">Object</span></span> | <span data-ttu-id="9e26b-624">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="9e26b-624">&lt;optional&gt;</span></span> | <span data-ttu-id="9e26b-625">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="9e26b-625">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="9e26b-626">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="9e26b-626">Object</span></span> | <span data-ttu-id="9e26b-627">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="9e26b-627">&lt;optional&gt;</span></span> | <span data-ttu-id="9e26b-628">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="9e26b-628">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="9e26b-629">function</span><span class="sxs-lookup"><span data-stu-id="9e26b-629">function</span></span>| <span data-ttu-id="9e26b-630">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="9e26b-630">&lt;optional&gt;</span></span>|<span data-ttu-id="9e26b-631">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="9e26b-631">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9e26b-632">要件</span><span class="sxs-lookup"><span data-stu-id="9e26b-632">Requirements</span></span>

|<span data-ttu-id="9e26b-633">要件</span><span class="sxs-lookup"><span data-stu-id="9e26b-633">Requirement</span></span>| <span data-ttu-id="9e26b-634">値</span><span class="sxs-lookup"><span data-stu-id="9e26b-634">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e26b-635">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="9e26b-635">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e26b-636">1.5</span><span class="sxs-lookup"><span data-stu-id="9e26b-636">1.5</span></span> |
|[<span data-ttu-id="9e26b-637">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="9e26b-637">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e26b-638">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9e26b-638">ReadItem</span></span> |
|[<span data-ttu-id="9e26b-639">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="9e26b-639">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9e26b-640">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="9e26b-640">Compose or Read</span></span>|
