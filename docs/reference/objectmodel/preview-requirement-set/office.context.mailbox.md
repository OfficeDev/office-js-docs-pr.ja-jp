---
title: Office のメールボックス-プレビュー要件セット
description: ''
ms.date: 10/21/2019
localization_priority: Normal
ms.openlocfilehash: 29922c9e05cc0380f1e54a16f3350c578d9e4cee
ms.sourcegitcommit: 499bf49b41205f8034c501d4db5fe4b02dab205e
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/22/2019
ms.locfileid: "37627070"
---
# <a name="mailbox"></a><span data-ttu-id="71b93-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="71b93-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="71b93-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="71b93-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="71b93-104">Microsoft Outlook の Outlook アドイン オブジェクト モデルへのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="71b93-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="71b93-105">要件</span><span class="sxs-lookup"><span data-stu-id="71b93-105">Requirements</span></span>

|<span data-ttu-id="71b93-106">要件</span><span class="sxs-lookup"><span data-stu-id="71b93-106">Requirement</span></span>| <span data-ttu-id="71b93-107">値</span><span class="sxs-lookup"><span data-stu-id="71b93-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="71b93-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="71b93-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="71b93-109">1.0</span><span class="sxs-lookup"><span data-stu-id="71b93-109">1.0</span></span>|
|[<span data-ttu-id="71b93-110">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="71b93-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="71b93-111">制限あり</span><span class="sxs-lookup"><span data-stu-id="71b93-111">Restricted</span></span>|
|[<span data-ttu-id="71b93-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="71b93-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="71b93-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="71b93-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="71b93-114">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="71b93-114">Members and methods</span></span>

| <span data-ttu-id="71b93-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="71b93-115">Member</span></span> | <span data-ttu-id="71b93-116">種類</span><span class="sxs-lookup"><span data-stu-id="71b93-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="71b93-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="71b93-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="71b93-118">メンバー</span><span class="sxs-lookup"><span data-stu-id="71b93-118">Member</span></span> |
| [<span data-ttu-id="71b93-119">masterCategories</span><span class="sxs-lookup"><span data-stu-id="71b93-119">masterCategories</span></span>](#mastercategories-mastercategories) | <span data-ttu-id="71b93-120">メンバー</span><span class="sxs-lookup"><span data-stu-id="71b93-120">Member</span></span> |
| [<span data-ttu-id="71b93-121">restUrl</span><span class="sxs-lookup"><span data-stu-id="71b93-121">restUrl</span></span>](#resturl-string) | <span data-ttu-id="71b93-122">メンバー</span><span class="sxs-lookup"><span data-stu-id="71b93-122">Member</span></span> |
| [<span data-ttu-id="71b93-123">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="71b93-123">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="71b93-124">メソッド</span><span class="sxs-lookup"><span data-stu-id="71b93-124">Method</span></span> |
| [<span data-ttu-id="71b93-125">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="71b93-125">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="71b93-126">メソッド</span><span class="sxs-lookup"><span data-stu-id="71b93-126">Method</span></span> |
| [<span data-ttu-id="71b93-127">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="71b93-127">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="71b93-128">メソッド</span><span class="sxs-lookup"><span data-stu-id="71b93-128">Method</span></span> |
| [<span data-ttu-id="71b93-129">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="71b93-129">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="71b93-130">メソッド</span><span class="sxs-lookup"><span data-stu-id="71b93-130">Method</span></span> |
| [<span data-ttu-id="71b93-131">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="71b93-131">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="71b93-132">メソッド</span><span class="sxs-lookup"><span data-stu-id="71b93-132">Method</span></span> |
| [<span data-ttu-id="71b93-133">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="71b93-133">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="71b93-134">メソッド</span><span class="sxs-lookup"><span data-stu-id="71b93-134">Method</span></span> |
| [<span data-ttu-id="71b93-135">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="71b93-135">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="71b93-136">メソッド</span><span class="sxs-lookup"><span data-stu-id="71b93-136">Method</span></span> |
| [<span data-ttu-id="71b93-137">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="71b93-137">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="71b93-138">メソッド</span><span class="sxs-lookup"><span data-stu-id="71b93-138">Method</span></span> |
| [<span data-ttu-id="71b93-139">displayNewMessageForm</span><span class="sxs-lookup"><span data-stu-id="71b93-139">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="71b93-140">メソッド</span><span class="sxs-lookup"><span data-stu-id="71b93-140">Method</span></span> |
| [<span data-ttu-id="71b93-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="71b93-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="71b93-142">メソッド</span><span class="sxs-lookup"><span data-stu-id="71b93-142">Method</span></span> |
| [<span data-ttu-id="71b93-143">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="71b93-143">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="71b93-144">メソッド</span><span class="sxs-lookup"><span data-stu-id="71b93-144">Method</span></span> |
| [<span data-ttu-id="71b93-145">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="71b93-145">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="71b93-146">メソッド</span><span class="sxs-lookup"><span data-stu-id="71b93-146">Method</span></span> |
| [<span data-ttu-id="71b93-147">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="71b93-147">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="71b93-148">メソッド</span><span class="sxs-lookup"><span data-stu-id="71b93-148">Method</span></span> |
| [<span data-ttu-id="71b93-149">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="71b93-149">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="71b93-150">メソッド</span><span class="sxs-lookup"><span data-stu-id="71b93-150">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="71b93-151">名前空間</span><span class="sxs-lookup"><span data-stu-id="71b93-151">Namespaces</span></span>

<span data-ttu-id="71b93-152">[diagnostics](Office.context.mailbox.diagnostics.md):Outlook アドインに診断情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="71b93-152">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="71b93-153">[item](Office.context.mailbox.item.md):Outlook アドインのメッセージや予定にアクセスするためのメソッドとプロパティを提供します。</span><span class="sxs-lookup"><span data-stu-id="71b93-153">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="71b93-154">[userProfile](Office.context.mailbox.userProfile.md):Outlook アドインのユーザーに関する情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="71b93-154">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="71b93-155">Members</span><span class="sxs-lookup"><span data-stu-id="71b93-155">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="71b93-156">ewsUrl: String</span><span class="sxs-lookup"><span data-stu-id="71b93-156">ewsUrl: String</span></span>

<span data-ttu-id="71b93-p101">このメール アカウントの Exchange Web サービス (EWS) エンドポイントの URL を取得します。読み取りモードのみです。</span><span class="sxs-lookup"><span data-stu-id="71b93-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="71b93-159">このメンバーは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="71b93-159">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="71b93-p102">
  `ewsUrl\` 値は、リモート サービスで、ユーザーのメールボックスに EWS 呼び出しを行うために使うことができます。たとえば、[選択したアイテムから添付ファイルを取得する](/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="71b93-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="71b93-162">アプリが閲覧モードで `ewsUrl` メンバーを呼び出すには、アプリのマニフェスト内に **ReadItem** アクセス許可が指定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="71b93-162">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="71b93-p103">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出してから、`ewsUrl` メンバーを使用する必要があります。アプリには、`saveAsync` メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="71b93-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="71b93-165">型</span><span class="sxs-lookup"><span data-stu-id="71b93-165">Type</span></span>

*   <span data-ttu-id="71b93-166">String</span><span class="sxs-lookup"><span data-stu-id="71b93-166">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="71b93-167">要件</span><span class="sxs-lookup"><span data-stu-id="71b93-167">Requirements</span></span>

|<span data-ttu-id="71b93-168">要件</span><span class="sxs-lookup"><span data-stu-id="71b93-168">Requirement</span></span>| <span data-ttu-id="71b93-169">値</span><span class="sxs-lookup"><span data-stu-id="71b93-169">Value</span></span>|
|---|---|
|[<span data-ttu-id="71b93-170">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="71b93-170">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="71b93-171">1.0</span><span class="sxs-lookup"><span data-stu-id="71b93-171">1.0</span></span>|
|[<span data-ttu-id="71b93-172">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="71b93-172">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="71b93-173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="71b93-173">ReadItem</span></span>|
|[<span data-ttu-id="71b93-174">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="71b93-174">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="71b93-175">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="71b93-175">Compose or Read</span></span>|

<br>

---
---

#### <a name="mastercategories-mastercategoriesjavascriptapioutlookofficemastercategories"></a><span data-ttu-id="71b93-176">masterCategories: [Mastercategories](/javascript/api/outlook/office.mastercategories)</span><span class="sxs-lookup"><span data-stu-id="71b93-176">masterCategories: [MasterCategories](/javascript/api/outlook/office.mastercategories)</span></span>

<span data-ttu-id="71b93-177">このメールボックスのカテゴリマスターリストを管理するためのメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="71b93-177">Gets an object that provides methods to manage the categories master list on this mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="71b93-178">このメンバーは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="71b93-178">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="71b93-179">種類</span><span class="sxs-lookup"><span data-stu-id="71b93-179">Type</span></span>

*   [<span data-ttu-id="71b93-180">MasterCategories</span><span class="sxs-lookup"><span data-stu-id="71b93-180">MasterCategories</span></span>](/javascript/api/outlook/office.mastercategories)

##### <a name="requirements"></a><span data-ttu-id="71b93-181">要件</span><span class="sxs-lookup"><span data-stu-id="71b93-181">Requirements</span></span>

|<span data-ttu-id="71b93-182">要件</span><span class="sxs-lookup"><span data-stu-id="71b93-182">Requirement</span></span>| <span data-ttu-id="71b93-183">値</span><span class="sxs-lookup"><span data-stu-id="71b93-183">Value</span></span>|
|---|---|
|[<span data-ttu-id="71b93-184">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="71b93-184">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="71b93-185">プレビュー</span><span class="sxs-lookup"><span data-stu-id="71b93-185">Preview</span></span> |
|[<span data-ttu-id="71b93-186">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="71b93-186">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="71b93-187">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="71b93-187">ReadWriteMailbox</span></span> |
|[<span data-ttu-id="71b93-188">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="71b93-188">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="71b93-189">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="71b93-189">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="71b93-190">例</span><span class="sxs-lookup"><span data-stu-id="71b93-190">Example</span></span>

<span data-ttu-id="71b93-191">この例では、このメールボックスのカテゴリマスターリストを取得します。</span><span class="sxs-lookup"><span data-stu-id="71b93-191">This example gets the categories master list for this mailbox.</span></span>

```js
Office.context.mailbox.masterCategories.getAsync(function (asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    console.log("Action failed with error: " + asyncResult.error.message);
  } else {
    console.log("Master categories: " + JSON.stringify(asyncResult.value));
  }
});
```

<br>

---
---

#### <a name="resturl-string"></a><span data-ttu-id="71b93-192">restUrl: String</span><span class="sxs-lookup"><span data-stu-id="71b93-192">restUrl: String</span></span>

<span data-ttu-id="71b93-193">この電子メール アカウントの REST エンドポイントの URL を取得します。</span><span class="sxs-lookup"><span data-stu-id="71b93-193">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="71b93-194">`restUrl` 値は、ユーザーのメールボックスに [REST API](/outlook/rest/) 呼び出しを行うために使用できます。</span><span class="sxs-lookup"><span data-stu-id="71b93-194">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="71b93-195">アプリが閲覧モードで `restUrl` メンバーを呼び出すには、アプリのマニフェスト内に **ReadItem** アクセス許可が指定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="71b93-195">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="71b93-p104">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出してから、`restUrl` メンバーを使用する必要があります。アプリには、`saveAsync` メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="71b93-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="71b93-198">型</span><span class="sxs-lookup"><span data-stu-id="71b93-198">Type</span></span>

*   <span data-ttu-id="71b93-199">String</span><span class="sxs-lookup"><span data-stu-id="71b93-199">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="71b93-200">要件</span><span class="sxs-lookup"><span data-stu-id="71b93-200">Requirements</span></span>

|<span data-ttu-id="71b93-201">要件</span><span class="sxs-lookup"><span data-stu-id="71b93-201">Requirement</span></span>| <span data-ttu-id="71b93-202">値</span><span class="sxs-lookup"><span data-stu-id="71b93-202">Value</span></span>|
|---|---|
|[<span data-ttu-id="71b93-203">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="71b93-203">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="71b93-204">1.5</span><span class="sxs-lookup"><span data-stu-id="71b93-204">1.5</span></span> |
|[<span data-ttu-id="71b93-205">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="71b93-205">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="71b93-206">ReadItem</span><span class="sxs-lookup"><span data-stu-id="71b93-206">ReadItem</span></span>|
|[<span data-ttu-id="71b93-207">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="71b93-207">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="71b93-208">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="71b93-208">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="71b93-209">メソッド</span><span class="sxs-lookup"><span data-stu-id="71b93-209">Methods</span></span>

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="71b93-210">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="71b93-210">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="71b93-211">サポートされているイベントのイベント ハンドラーを追加します。</span><span class="sxs-lookup"><span data-stu-id="71b93-211">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="71b93-212">現在、サポートされている`Office.EventType.ItemChanged`イベント`Office.EventType.OfficeThemeChanged`の種類はおよびです。</span><span class="sxs-lookup"><span data-stu-id="71b93-212">Currently, the supported event types are `Office.EventType.ItemChanged` and `Office.EventType.OfficeThemeChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="71b93-213">パラメーター</span><span class="sxs-lookup"><span data-stu-id="71b93-213">Parameters</span></span>

| <span data-ttu-id="71b93-214">名前</span><span class="sxs-lookup"><span data-stu-id="71b93-214">Name</span></span> | <span data-ttu-id="71b93-215">種類</span><span class="sxs-lookup"><span data-stu-id="71b93-215">Type</span></span> | <span data-ttu-id="71b93-216">属性</span><span class="sxs-lookup"><span data-stu-id="71b93-216">Attributes</span></span> | <span data-ttu-id="71b93-217">説明</span><span class="sxs-lookup"><span data-stu-id="71b93-217">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="71b93-218">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="71b93-218">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="71b93-219">ハンドラーを呼び出す必要のあるイベント。</span><span class="sxs-lookup"><span data-stu-id="71b93-219">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="71b93-220">Function</span><span class="sxs-lookup"><span data-stu-id="71b93-220">Function</span></span> || <span data-ttu-id="71b93-p105">イベントを処理する関数。関数は、オブジェクト リテラルである単一パラメーターを受け入れる必要があります。パラメーターの `type` プロパティは、`addHandlerAsync` に渡される `eventType` パラメーターと一致します。</span><span class="sxs-lookup"><span data-stu-id="71b93-p105">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="71b93-224">Object</span><span class="sxs-lookup"><span data-stu-id="71b93-224">Object</span></span> | <span data-ttu-id="71b93-225">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="71b93-225">&lt;optional&gt;</span></span> | <span data-ttu-id="71b93-226">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="71b93-226">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="71b93-227">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="71b93-227">Object</span></span> | <span data-ttu-id="71b93-228">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="71b93-228">&lt;optional&gt;</span></span> | <span data-ttu-id="71b93-229">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="71b93-229">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="71b93-230">function</span><span class="sxs-lookup"><span data-stu-id="71b93-230">function</span></span>| <span data-ttu-id="71b93-231">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="71b93-231">&lt;optional&gt;</span></span>|<span data-ttu-id="71b93-232">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="71b93-232">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="71b93-233">要件</span><span class="sxs-lookup"><span data-stu-id="71b93-233">Requirements</span></span>

|<span data-ttu-id="71b93-234">要件</span><span class="sxs-lookup"><span data-stu-id="71b93-234">Requirement</span></span>| <span data-ttu-id="71b93-235">値</span><span class="sxs-lookup"><span data-stu-id="71b93-235">Value</span></span>|
|---|---|
|[<span data-ttu-id="71b93-236">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="71b93-236">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="71b93-237">1.5</span><span class="sxs-lookup"><span data-stu-id="71b93-237">1.5</span></span> |
|[<span data-ttu-id="71b93-238">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="71b93-238">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="71b93-239">ReadItem</span><span class="sxs-lookup"><span data-stu-id="71b93-239">ReadItem</span></span> |
|[<span data-ttu-id="71b93-240">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="71b93-240">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="71b93-241">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="71b93-241">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="71b93-242">例</span><span class="sxs-lookup"><span data-stu-id="71b93-242">Example</span></span>

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
}
```

<br>

---
---

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="71b93-243">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="71b93-243">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="71b93-244">REST 形式のアイテム ID を EWS 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="71b93-244">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="71b93-245">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="71b93-245">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="71b93-p106">REST API ([Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) や [Microsoft Graph](https://graph.microsoft.io/) など) で取得されたアイテム ID は、Exchange Web サービス (EWS) に使用される形式とは異なる形式を使用します。`convertToEwsId` メソッドは、REST 形式の ID を EWS 用の適切な形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="71b93-p106">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="71b93-248">パラメーター</span><span class="sxs-lookup"><span data-stu-id="71b93-248">Parameters</span></span>

|<span data-ttu-id="71b93-249">名前</span><span class="sxs-lookup"><span data-stu-id="71b93-249">Name</span></span>| <span data-ttu-id="71b93-250">種類</span><span class="sxs-lookup"><span data-stu-id="71b93-250">Type</span></span>| <span data-ttu-id="71b93-251">説明</span><span class="sxs-lookup"><span data-stu-id="71b93-251">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="71b93-252">String</span><span class="sxs-lookup"><span data-stu-id="71b93-252">String</span></span>|<span data-ttu-id="71b93-253">Outlook REST API 形式のアイテム ID</span><span class="sxs-lookup"><span data-stu-id="71b93-253">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="71b93-254">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="71b93-254">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion)|<span data-ttu-id="71b93-255">アイテム ID の取得に使用された Outlook REST API のバージョンを示す値。</span><span class="sxs-lookup"><span data-stu-id="71b93-255">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="71b93-256">要件</span><span class="sxs-lookup"><span data-stu-id="71b93-256">Requirements</span></span>

|<span data-ttu-id="71b93-257">要件</span><span class="sxs-lookup"><span data-stu-id="71b93-257">Requirement</span></span>| <span data-ttu-id="71b93-258">値</span><span class="sxs-lookup"><span data-stu-id="71b93-258">Value</span></span>|
|---|---|
|[<span data-ttu-id="71b93-259">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="71b93-259">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="71b93-260">1.3</span><span class="sxs-lookup"><span data-stu-id="71b93-260">1.3</span></span>|
|[<span data-ttu-id="71b93-261">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="71b93-261">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="71b93-262">制限あり</span><span class="sxs-lookup"><span data-stu-id="71b93-262">Restricted</span></span>|
|[<span data-ttu-id="71b93-263">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="71b93-263">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="71b93-264">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="71b93-264">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="71b93-265">戻り値:</span><span class="sxs-lookup"><span data-stu-id="71b93-265">Returns:</span></span>

<span data-ttu-id="71b93-266">型:String</span><span class="sxs-lookup"><span data-stu-id="71b93-266">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="71b93-267">例</span><span class="sxs-lookup"><span data-stu-id="71b93-267">Example</span></span>

```js
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime"></a><span data-ttu-id="71b93-268">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="71b93-268">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime)}</span></span>

<span data-ttu-id="71b93-269">クライアントのローカル時間で時間情報が含まれている辞書を取得します。</span><span class="sxs-lookup"><span data-stu-id="71b93-269">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="71b93-p107">Outlook on the web または Outlook デスクトップのメールアプリでは、日付と時刻に異なるタイムゾーンを使用できます。Outlook デスクトップは、クライアント コンピューターのタイム ゾーンを使用します。Outlook on the web は、Exchange 管理センター (EAC) で設定されたタイム ゾーンを使用します。日付と時刻の値は、ユーザー インターフェイスに表示される値が、常にユーザーが期待するタイム ゾーンと一致するように処理する必要があります。</span><span class="sxs-lookup"><span data-stu-id="71b93-p107">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times. Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="71b93-p108">Outlook デスクトップ クライアントでメール アプリを実行している場合、`convertToLocalClientTime` メソッドは、クライアント コンピューターのタイム ゾーンに設定された値のディクショナリ オブジェクトを返します。Outlook on the web でメール アプリを実行している場合、`convertToLocalClientTime` メソッドは、EAC で指定したタイム ゾーンに設定された値のディクショナリ オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="71b93-p108">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="71b93-275">パラメーター</span><span class="sxs-lookup"><span data-stu-id="71b93-275">Parameters</span></span>

|<span data-ttu-id="71b93-276">名前</span><span class="sxs-lookup"><span data-stu-id="71b93-276">Name</span></span>| <span data-ttu-id="71b93-277">種類</span><span class="sxs-lookup"><span data-stu-id="71b93-277">Type</span></span>| <span data-ttu-id="71b93-278">説明</span><span class="sxs-lookup"><span data-stu-id="71b93-278">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="71b93-279">日付</span><span class="sxs-lookup"><span data-stu-id="71b93-279">Date</span></span>|<span data-ttu-id="71b93-280">日付オブジェクト</span><span class="sxs-lookup"><span data-stu-id="71b93-280">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="71b93-281">要件</span><span class="sxs-lookup"><span data-stu-id="71b93-281">Requirements</span></span>

|<span data-ttu-id="71b93-282">要件</span><span class="sxs-lookup"><span data-stu-id="71b93-282">Requirement</span></span>| <span data-ttu-id="71b93-283">値</span><span class="sxs-lookup"><span data-stu-id="71b93-283">Value</span></span>|
|---|---|
|[<span data-ttu-id="71b93-284">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="71b93-284">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="71b93-285">1.0</span><span class="sxs-lookup"><span data-stu-id="71b93-285">1.0</span></span>|
|[<span data-ttu-id="71b93-286">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="71b93-286">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="71b93-287">ReadItem</span><span class="sxs-lookup"><span data-stu-id="71b93-287">ReadItem</span></span>|
|[<span data-ttu-id="71b93-288">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="71b93-288">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="71b93-289">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="71b93-289">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="71b93-290">戻り値:</span><span class="sxs-lookup"><span data-stu-id="71b93-290">Returns:</span></span>

<span data-ttu-id="71b93-291">型:[LocalClientTime](/javascript/api/outlook/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="71b93-291">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime)</span></span>

<br>

---
---

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="71b93-292">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="71b93-292">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="71b93-293">EWS 形式のアイテム ID を REST 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="71b93-293">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="71b93-294">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="71b93-294">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="71b93-p109">EWS または `itemId` プロパティで取得されるアイテム ID は、REST API ([Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) や [Microsoft Graph](https://graph.microsoft.io/) など) に使用される形式とは異なる形式を使用します。`convertToRestId` メソッドは、EWS 形式の ID を REST 用の適切な形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="71b93-p109">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="71b93-297">パラメーター</span><span class="sxs-lookup"><span data-stu-id="71b93-297">Parameters</span></span>

|<span data-ttu-id="71b93-298">名前</span><span class="sxs-lookup"><span data-stu-id="71b93-298">Name</span></span>| <span data-ttu-id="71b93-299">種類</span><span class="sxs-lookup"><span data-stu-id="71b93-299">Type</span></span>| <span data-ttu-id="71b93-300">説明</span><span class="sxs-lookup"><span data-stu-id="71b93-300">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="71b93-301">String</span><span class="sxs-lookup"><span data-stu-id="71b93-301">String</span></span>|<span data-ttu-id="71b93-302">Exchange Web サービス (EWS) 形式のアイテム ID</span><span class="sxs-lookup"><span data-stu-id="71b93-302">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="71b93-303">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="71b93-303">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion)|<span data-ttu-id="71b93-304">変換後の ID を使用する Outlook REST API のバージョンを示す値。</span><span class="sxs-lookup"><span data-stu-id="71b93-304">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="71b93-305">要件</span><span class="sxs-lookup"><span data-stu-id="71b93-305">Requirements</span></span>

|<span data-ttu-id="71b93-306">要件</span><span class="sxs-lookup"><span data-stu-id="71b93-306">Requirement</span></span>| <span data-ttu-id="71b93-307">値</span><span class="sxs-lookup"><span data-stu-id="71b93-307">Value</span></span>|
|---|---|
|[<span data-ttu-id="71b93-308">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="71b93-308">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="71b93-309">1.3</span><span class="sxs-lookup"><span data-stu-id="71b93-309">1.3</span></span>|
|[<span data-ttu-id="71b93-310">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="71b93-310">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="71b93-311">制限あり</span><span class="sxs-lookup"><span data-stu-id="71b93-311">Restricted</span></span>|
|[<span data-ttu-id="71b93-312">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="71b93-312">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="71b93-313">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="71b93-313">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="71b93-314">戻り値:</span><span class="sxs-lookup"><span data-stu-id="71b93-314">Returns:</span></span>

<span data-ttu-id="71b93-315">型:String</span><span class="sxs-lookup"><span data-stu-id="71b93-315">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="71b93-316">例</span><span class="sxs-lookup"><span data-stu-id="71b93-316">Example</span></span>

```js
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="71b93-317">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="71b93-317">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="71b93-318">時間情報が含まれているディクショナリから日付オブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="71b93-318">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="71b93-319">`convertToUtcClientTime` メソッドは、ローカルの日付と時刻を含むディクショナリを、ローカルの日付と時刻の正しい値を持つ日付オブジェクトに変換します。</span><span class="sxs-lookup"><span data-stu-id="71b93-319">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="71b93-320">パラメーター</span><span class="sxs-lookup"><span data-stu-id="71b93-320">Parameters</span></span>

|<span data-ttu-id="71b93-321">名前</span><span class="sxs-lookup"><span data-stu-id="71b93-321">Name</span></span>| <span data-ttu-id="71b93-322">種類</span><span class="sxs-lookup"><span data-stu-id="71b93-322">Type</span></span>| <span data-ttu-id="71b93-323">説明</span><span class="sxs-lookup"><span data-stu-id="71b93-323">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="71b93-324">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="71b93-324">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime)|<span data-ttu-id="71b93-325">変換するローカル時刻の値。</span><span class="sxs-lookup"><span data-stu-id="71b93-325">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="71b93-326">要件</span><span class="sxs-lookup"><span data-stu-id="71b93-326">Requirements</span></span>

|<span data-ttu-id="71b93-327">要件</span><span class="sxs-lookup"><span data-stu-id="71b93-327">Requirement</span></span>| <span data-ttu-id="71b93-328">値</span><span class="sxs-lookup"><span data-stu-id="71b93-328">Value</span></span>|
|---|---|
|[<span data-ttu-id="71b93-329">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="71b93-329">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="71b93-330">1.0</span><span class="sxs-lookup"><span data-stu-id="71b93-330">1.0</span></span>|
|[<span data-ttu-id="71b93-331">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="71b93-331">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="71b93-332">ReadItem</span><span class="sxs-lookup"><span data-stu-id="71b93-332">ReadItem</span></span>|
|[<span data-ttu-id="71b93-333">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="71b93-333">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="71b93-334">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="71b93-334">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="71b93-335">戻り値:</span><span class="sxs-lookup"><span data-stu-id="71b93-335">Returns:</span></span>

<span data-ttu-id="71b93-336">時間が UTC で表現された日付オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="71b93-336">A Date object with the time expressed in UTC.</span></span>

<span data-ttu-id="71b93-337">型: Date</span><span class="sxs-lookup"><span data-stu-id="71b93-337">Type: Date</span></span>

##### <a name="example"></a><span data-ttu-id="71b93-338">例</span><span class="sxs-lookup"><span data-stu-id="71b93-338">Example</span></span>

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

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="71b93-339">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="71b93-339">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="71b93-340">既存の予定を表示します。</span><span class="sxs-lookup"><span data-stu-id="71b93-340">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="71b93-341">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="71b93-341">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="71b93-342">`displayAppointmentForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存の予定を開きます。</span><span class="sxs-lookup"><span data-stu-id="71b93-342">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="71b93-p110">Outlook on Mac では、このメソッドを使用して定期的な系列に含まれない単発の予定や定期的な系列のマスター予定を表示できます。ただし、系列のインスタンスは表示できません。これは、Outlook on Mac では定期的な系列のインスタンスのプロパティ (アイテム ID を含む) にアクセスできないためです。</span><span class="sxs-lookup"><span data-stu-id="71b93-p110">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="71b93-345">Outlook on the web では、このメソッドはフォームの本文が 32KB 以下の文字数の場合にのみ指定のフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="71b93-345">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="71b93-346">指定のアイテム識別子が既存の予定を表していない場合は、クライアント コンピューターまたはデバイスで空のウィンドウが開き、エラー メッセージは返されません。</span><span class="sxs-lookup"><span data-stu-id="71b93-346">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="71b93-347">パラメーター</span><span class="sxs-lookup"><span data-stu-id="71b93-347">Parameters</span></span>

|<span data-ttu-id="71b93-348">名前</span><span class="sxs-lookup"><span data-stu-id="71b93-348">Name</span></span>| <span data-ttu-id="71b93-349">種類</span><span class="sxs-lookup"><span data-stu-id="71b93-349">Type</span></span>| <span data-ttu-id="71b93-350">説明</span><span class="sxs-lookup"><span data-stu-id="71b93-350">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="71b93-351">String</span><span class="sxs-lookup"><span data-stu-id="71b93-351">String</span></span>|<span data-ttu-id="71b93-352">既存の予定の Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="71b93-352">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="71b93-353">要件</span><span class="sxs-lookup"><span data-stu-id="71b93-353">Requirements</span></span>

|<span data-ttu-id="71b93-354">要件</span><span class="sxs-lookup"><span data-stu-id="71b93-354">Requirement</span></span>| <span data-ttu-id="71b93-355">値</span><span class="sxs-lookup"><span data-stu-id="71b93-355">Value</span></span>|
|---|---|
|[<span data-ttu-id="71b93-356">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="71b93-356">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="71b93-357">1.0</span><span class="sxs-lookup"><span data-stu-id="71b93-357">1.0</span></span>|
|[<span data-ttu-id="71b93-358">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="71b93-358">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="71b93-359">ReadItem</span><span class="sxs-lookup"><span data-stu-id="71b93-359">ReadItem</span></span>|
|[<span data-ttu-id="71b93-360">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="71b93-360">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="71b93-361">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="71b93-361">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="71b93-362">例</span><span class="sxs-lookup"><span data-stu-id="71b93-362">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

<br>

---
---

#### <a name="displaymessageformitemid"></a><span data-ttu-id="71b93-363">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="71b93-363">displayMessageForm(itemId)</span></span>

<span data-ttu-id="71b93-364">既存のメッセージを表示します。</span><span class="sxs-lookup"><span data-stu-id="71b93-364">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="71b93-365">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="71b93-365">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="71b93-366">`displayMessageForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存のメッセージを開きます。</span><span class="sxs-lookup"><span data-stu-id="71b93-366">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="71b93-367">Outlook on the web では、このメソッドはフォームの本文が 32 KB 以下の文字数の場合にのみ指定のフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="71b93-367">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="71b93-368">指定のアイテム識別子が既存のメッセージを表していない場合は、クライアント コンピューターにはメッセージは表示されず、エラー メッセージも返されません。</span><span class="sxs-lookup"><span data-stu-id="71b93-368">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="71b93-p111">予定を表す `itemId` を含む `displayMessageForm` を使用しないでください。既存の予定を表示するには、`displayAppointmentForm` メソッドを使用します。新しい予定を作成するフォームを表示するには、`displayNewAppointmentForm` メソッドを使用します。</span><span class="sxs-lookup"><span data-stu-id="71b93-p111">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="71b93-371">パラメーター</span><span class="sxs-lookup"><span data-stu-id="71b93-371">Parameters</span></span>

|<span data-ttu-id="71b93-372">名前</span><span class="sxs-lookup"><span data-stu-id="71b93-372">Name</span></span>| <span data-ttu-id="71b93-373">種類</span><span class="sxs-lookup"><span data-stu-id="71b93-373">Type</span></span>| <span data-ttu-id="71b93-374">説明</span><span class="sxs-lookup"><span data-stu-id="71b93-374">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="71b93-375">String</span><span class="sxs-lookup"><span data-stu-id="71b93-375">String</span></span>|<span data-ttu-id="71b93-376">既存のメッセージの Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="71b93-376">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="71b93-377">要件</span><span class="sxs-lookup"><span data-stu-id="71b93-377">Requirements</span></span>

|<span data-ttu-id="71b93-378">要件</span><span class="sxs-lookup"><span data-stu-id="71b93-378">Requirement</span></span>| <span data-ttu-id="71b93-379">値</span><span class="sxs-lookup"><span data-stu-id="71b93-379">Value</span></span>|
|---|---|
|[<span data-ttu-id="71b93-380">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="71b93-380">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="71b93-381">1.0</span><span class="sxs-lookup"><span data-stu-id="71b93-381">1.0</span></span>|
|[<span data-ttu-id="71b93-382">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="71b93-382">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="71b93-383">ReadItem</span><span class="sxs-lookup"><span data-stu-id="71b93-383">ReadItem</span></span>|
|[<span data-ttu-id="71b93-384">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="71b93-384">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="71b93-385">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="71b93-385">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="71b93-386">例</span><span class="sxs-lookup"><span data-stu-id="71b93-386">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

<br>

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="71b93-387">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="71b93-387">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="71b93-388">新しい予定を作成するためのフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="71b93-388">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="71b93-389">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="71b93-389">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="71b93-p112">`displayNewAppointmentForm` メソッドを使用すると、ユーザーが新しい予定または会議を作成できるフォームが開きます。パラメーターを指定すると、予定のフォーム フィールドにパラメーターの内容が自動的に設定されます。</span><span class="sxs-lookup"><span data-stu-id="71b93-p112">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="71b93-p113">Outlook on the web およびモバイル デバイスでは、このメソッドは常に出席者フィールドが含まれるフォームを表示します。入力引数として出席者を指定しないと、このメソッドは **[保存]** ボタンのあるフォームを表示します。出席者を指定した場合には、フォームにその出席者と **[送信]** ボタンが表示されます。</span><span class="sxs-lookup"><span data-stu-id="71b93-p113">In Outlook on the web and mobile devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="71b93-p114">Outlook リッチ クライアントと Outlook RT で、`requiredAttendees`、`optionalAttendees`、または `resources` パラメーターに出席者またはリソースを指定し、このメソッドを実行すると、**[送信]** ボタンがある会議フォームが表示されます。受信者を指定せずにこのメソッドを実行すると、**[保存して閉じる]** ボタンがある予定フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="71b93-p114">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="71b93-397">パラメーターのいずれかが指定のサイズ制限を超える場合、または不明なパラメーター名が指定されている場合は、例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="71b93-397">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="71b93-398">パラメーター</span><span class="sxs-lookup"><span data-stu-id="71b93-398">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="71b93-399">すべてのパラメーターは省略可能です。</span><span class="sxs-lookup"><span data-stu-id="71b93-399">All parameters are optional.</span></span>

|<span data-ttu-id="71b93-400">名前</span><span class="sxs-lookup"><span data-stu-id="71b93-400">Name</span></span>| <span data-ttu-id="71b93-401">種類</span><span class="sxs-lookup"><span data-stu-id="71b93-401">Type</span></span>| <span data-ttu-id="71b93-402">説明</span><span class="sxs-lookup"><span data-stu-id="71b93-402">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="71b93-403">Object</span><span class="sxs-lookup"><span data-stu-id="71b93-403">Object</span></span> | <span data-ttu-id="71b93-404">新しい予定を記述するパラメーターのディクショナリ。</span><span class="sxs-lookup"><span data-stu-id="71b93-404">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="71b93-405">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="71b93-405">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="71b93-p115">予定に必要な各出席者について、メール アドレスを含む文字列の配列、または `EmailAddressDetails` オブジェクトを含む配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="71b93-p115">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="71b93-408">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="71b93-408">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="71b93-p116">予定の各任意出席者について、メール アドレスを含む文字列の配列、または `EmailAddressDetails` オブジェクトを含む配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="71b93-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="71b93-411">日付</span><span class="sxs-lookup"><span data-stu-id="71b93-411">Date</span></span> | <span data-ttu-id="71b93-412">予定の開始日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="71b93-412">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="71b93-413">日付</span><span class="sxs-lookup"><span data-stu-id="71b93-413">Date</span></span> | <span data-ttu-id="71b93-414">予定の終了日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="71b93-414">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="71b93-415">String</span><span class="sxs-lookup"><span data-stu-id="71b93-415">String</span></span> | <span data-ttu-id="71b93-p117">予定の場所を含む文字列。文字列は最大 255 文字に制限されます。</span><span class="sxs-lookup"><span data-stu-id="71b93-p117">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="71b93-418">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="71b93-418">Array.&lt;String&gt;</span></span> | <span data-ttu-id="71b93-p118">予定に必要なリソースを含む文字列の配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="71b93-p118">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="71b93-421">String</span><span class="sxs-lookup"><span data-stu-id="71b93-421">String</span></span> | <span data-ttu-id="71b93-p119">予定の件名を含む文字列です。文字列は最大 255 文字に制限されます。</span><span class="sxs-lookup"><span data-stu-id="71b93-p119">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="71b93-424">String</span><span class="sxs-lookup"><span data-stu-id="71b93-424">String</span></span> | <span data-ttu-id="71b93-p120">予定の本文。本文の内容は、最大サイズが 32 KB に制限されます。</span><span class="sxs-lookup"><span data-stu-id="71b93-p120">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="71b93-427">要件</span><span class="sxs-lookup"><span data-stu-id="71b93-427">Requirements</span></span>

|<span data-ttu-id="71b93-428">要件</span><span class="sxs-lookup"><span data-stu-id="71b93-428">Requirement</span></span>| <span data-ttu-id="71b93-429">値</span><span class="sxs-lookup"><span data-stu-id="71b93-429">Value</span></span>|
|---|---|
|[<span data-ttu-id="71b93-430">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="71b93-430">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="71b93-431">1.0</span><span class="sxs-lookup"><span data-stu-id="71b93-431">1.0</span></span>|
|[<span data-ttu-id="71b93-432">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="71b93-432">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="71b93-433">ReadItem</span><span class="sxs-lookup"><span data-stu-id="71b93-433">ReadItem</span></span>|
|[<span data-ttu-id="71b93-434">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="71b93-434">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="71b93-435">読み取り</span><span class="sxs-lookup"><span data-stu-id="71b93-435">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="71b93-436">例</span><span class="sxs-lookup"><span data-stu-id="71b93-436">Example</span></span>

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

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="71b93-437">displayNewMessageForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="71b93-437">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="71b93-438">新しいメッセージを作成するためのフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="71b93-438">Displays a form for creating a new message.</span></span>

<span data-ttu-id="71b93-p121">The `displayNewMessageForm` method opens a form that enables the user to create a new message. If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span><span class="sxs-lookup"><span data-stu-id="71b93-p121">The `displayNewMessageForm` method opens a form that enables the user to create a new message. If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="71b93-441">パラメーターのいずれかが指定のサイズ制限を超える場合、または不明なパラメーター名が指定されている場合は、例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="71b93-441">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="71b93-442">パラメーター</span><span class="sxs-lookup"><span data-stu-id="71b93-442">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="71b93-443">すべてのパラメーターは省略可能です。</span><span class="sxs-lookup"><span data-stu-id="71b93-443">All parameters are optional.</span></span>

|<span data-ttu-id="71b93-444">名前</span><span class="sxs-lookup"><span data-stu-id="71b93-444">Name</span></span>| <span data-ttu-id="71b93-445">種類</span><span class="sxs-lookup"><span data-stu-id="71b93-445">Type</span></span>| <span data-ttu-id="71b93-446">説明</span><span class="sxs-lookup"><span data-stu-id="71b93-446">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="71b93-447">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="71b93-447">Object</span></span> | <span data-ttu-id="71b93-448">新しいメッセージを記述するパラメータの辞書。</span><span class="sxs-lookup"><span data-stu-id="71b93-448">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="71b93-449">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="71b93-449">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="71b93-p122">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line. The array is limited to a maximum of 100 entries.</span><span class="sxs-lookup"><span data-stu-id="71b93-p122">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="71b93-452">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="71b93-452">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="71b93-p123">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line. The array is limited to a maximum of 100 entries.</span><span class="sxs-lookup"><span data-stu-id="71b93-p123">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="71b93-455">配列。&lt;文字列&gt; | 配列。&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="71b93-455">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="71b93-p124">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line. The array is limited to a maximum of 100 entries.</span><span class="sxs-lookup"><span data-stu-id="71b93-p124">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="71b93-458">String</span><span class="sxs-lookup"><span data-stu-id="71b93-458">String</span></span> | <span data-ttu-id="71b93-p125">A string containing the subject of the message. The string is limited to a maximum of 255 characters.</span><span class="sxs-lookup"><span data-stu-id="71b93-p125">A string containing the subject of the message. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="71b93-461">String</span><span class="sxs-lookup"><span data-stu-id="71b93-461">String</span></span> | <span data-ttu-id="71b93-p126">The HTML body of the message. The body content is limited to a maximum size of 32 KB.</span><span class="sxs-lookup"><span data-stu-id="71b93-p126">The HTML body of the message. The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="71b93-464">配列。&lt;オブジェクト&gt;</span><span class="sxs-lookup"><span data-stu-id="71b93-464">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="71b93-465">添付ファイルまたは添付アイテムである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="71b93-465">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="71b93-466">String</span><span class="sxs-lookup"><span data-stu-id="71b93-466">String</span></span> | <span data-ttu-id="71b93-p127">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="71b93-p127">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="71b93-469">String</span><span class="sxs-lookup"><span data-stu-id="71b93-469">String</span></span> | <span data-ttu-id="71b93-470">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="71b93-470">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="71b93-471">文字列</span><span class="sxs-lookup"><span data-stu-id="71b93-471">String</span></span> | <span data-ttu-id="71b93-p128">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="71b93-p128">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="71b93-474">ブール値</span><span class="sxs-lookup"><span data-stu-id="71b93-474">Boolean</span></span> | <span data-ttu-id="71b93-p129">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="71b93-p129">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="71b93-477">String</span><span class="sxs-lookup"><span data-stu-id="71b93-477">String</span></span> | <span data-ttu-id="71b93-p130">Only used if `type` is set to `item`. The EWS item id of the existing e-mail you want to attach to the new message. This is a string up to 100 characters.</span><span class="sxs-lookup"><span data-stu-id="71b93-p130">Only used if `type` is set to `item`. The EWS item id of the existing e-mail you want to attach to the new message. This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="71b93-481">要件</span><span class="sxs-lookup"><span data-stu-id="71b93-481">Requirements</span></span>

|<span data-ttu-id="71b93-482">要件</span><span class="sxs-lookup"><span data-stu-id="71b93-482">Requirement</span></span>| <span data-ttu-id="71b93-483">値</span><span class="sxs-lookup"><span data-stu-id="71b93-483">Value</span></span>|
|---|---|
|[<span data-ttu-id="71b93-484">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="71b93-484">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="71b93-485">1.6</span><span class="sxs-lookup"><span data-stu-id="71b93-485">1.6</span></span> |
|[<span data-ttu-id="71b93-486">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="71b93-486">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="71b93-487">ReadItem</span><span class="sxs-lookup"><span data-stu-id="71b93-487">ReadItem</span></span>|
|[<span data-ttu-id="71b93-488">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="71b93-488">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="71b93-489">読み取り</span><span class="sxs-lookup"><span data-stu-id="71b93-489">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="71b93-490">例</span><span class="sxs-lookup"><span data-stu-id="71b93-490">Example</span></span>

```js
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

<br>

---
---

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="71b93-491">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="71b93-491">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="71b93-492">REST API または Exchange Web サービスを呼び出すために使用するトークンを含む文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="71b93-492">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="71b93-p131">`getCallbackTokenAsync` メソッドは、ユーザーのメールボックスをホストする Exchange Server から不透明なトークンを取得する非同期の呼び出しを行います。コールバック トークンの有効期間は 5 分です。</span><span class="sxs-lookup"><span data-stu-id="71b93-p131">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="71b93-495">可能であれば、アドインでは Exchange Web サービスの代わりに REST API を使用することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="71b93-495">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span>

<span data-ttu-id="71b93-496">読み取りモード`getCallbackTokenAsync`でメソッドを呼び出すには、 **ReadItem**の最低限のアクセス許可レベルが必要です。</span><span class="sxs-lookup"><span data-stu-id="71b93-496">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="71b93-497">新規`getCallbackTokenAsync`作成モードで呼び出しを行うには、アイテムを保存しておく必要があります。</span><span class="sxs-lookup"><span data-stu-id="71b93-497">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="71b93-498">この[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback)メソッドには、 **readwriteitem**の最小アクセス許可レベルが必要です。</span><span class="sxs-lookup"><span data-stu-id="71b93-498">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

<span data-ttu-id="71b93-499">**REST トークン**</span><span class="sxs-lookup"><span data-stu-id="71b93-499">**REST Tokens**</span></span>

<span data-ttu-id="71b93-p133">REST トークンが要求された場合 (`options.isRest = true`)、結果トークンは Exchange Web サービスの呼び出しを認証するためには機能しません。アドインがマニフェストで [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) アクセス許可を指定していない限り、トークンの範囲は現在のアイテムとその添付ファイルへの読み取り専用アクセスに制限されます。`ReadWriteMailbox` アクセス許可が指定されている場合は、結果トークンは、メールを送信する機能など、メール、カレンダー、連絡先への読み取り/書き込みアクセスを付与します。</span><span class="sxs-lookup"><span data-stu-id="71b93-p133">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="71b93-503">アドインでは、`restUrl` プロパティを使用して、REST API 呼び出しを行うときに使用する正しい URL を決定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="71b93-503">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="71b93-504">**EWS トークン**</span><span class="sxs-lookup"><span data-stu-id="71b93-504">**EWS Tokens**</span></span>

<span data-ttu-id="71b93-p134">EWS トークンが要求された場合 (`options.isRest = false`)、結果トークンは REST API 呼び出しを認証するためには機能しません。トークンの範囲は、現在のアイテムへのアクセスに制限されます。</span><span class="sxs-lookup"><span data-stu-id="71b93-p134">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="71b93-507">アドインでは、`ewsUrl` プロパティを使用して、EWS 呼び出しを行うときに使用する正しい URL を決定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="71b93-507">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

<span data-ttu-id="71b93-508">トークンと、添付ファイル識別子またはアイテム識別子の両方をサードパーティのシステムに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="71b93-508">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="71b93-509">サードパーティシステムは、トークンをベアラー認証トークンとして使用して、Exchange Web サービス (EWS) の[Getattachment](/exchange/client-developer/web-service-reference/getattachment-operation)操作または[GetItem](/exchange/client-developer/web-service-reference/getitem-operation)操作を呼び出して、添付ファイルまたはアイテムを取得します。</span><span class="sxs-lookup"><span data-stu-id="71b93-509">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to retrieve an attachment or item.</span></span> <span data-ttu-id="71b93-510">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="71b93-510">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

##### <a name="parameters"></a><span data-ttu-id="71b93-511">パラメーター</span><span class="sxs-lookup"><span data-stu-id="71b93-511">Parameters</span></span>

|<span data-ttu-id="71b93-512">名前</span><span class="sxs-lookup"><span data-stu-id="71b93-512">Name</span></span>| <span data-ttu-id="71b93-513">種類</span><span class="sxs-lookup"><span data-stu-id="71b93-513">Type</span></span>| <span data-ttu-id="71b93-514">属性</span><span class="sxs-lookup"><span data-stu-id="71b93-514">Attributes</span></span>| <span data-ttu-id="71b93-515">説明</span><span class="sxs-lookup"><span data-stu-id="71b93-515">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="71b93-516">Object</span><span class="sxs-lookup"><span data-stu-id="71b93-516">Object</span></span> | <span data-ttu-id="71b93-517">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="71b93-517">&lt;optional&gt;</span></span> | <span data-ttu-id="71b93-518">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="71b93-518">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="71b93-519">Boolean</span><span class="sxs-lookup"><span data-stu-id="71b93-519">Boolean</span></span> |  <span data-ttu-id="71b93-520">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="71b93-520">&lt;optional&gt;</span></span> | <span data-ttu-id="71b93-p136">提供されたトークンを Outlook REST API または Exchange Web サービスに使用するかどうかを決定します。既定値は、`false` です。</span><span class="sxs-lookup"><span data-stu-id="71b93-p136">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="71b93-523">Object</span><span class="sxs-lookup"><span data-stu-id="71b93-523">Object</span></span> |  <span data-ttu-id="71b93-524">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="71b93-524">&lt;optional&gt;</span></span> | <span data-ttu-id="71b93-525">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="71b93-525">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="71b93-526">function</span><span class="sxs-lookup"><span data-stu-id="71b93-526">function</span></span>||<span data-ttu-id="71b93-527">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="71b93-527">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="71b93-528">トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="71b93-528">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="71b93-529">エラーが発生した場合、 `asyncResult.error` および `asyncResult.diagnostics` のプロパティで追加情報が提供される場合があります。</span><span class="sxs-lookup"><span data-stu-id="71b93-529">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|

##### <a name="errors"></a><span data-ttu-id="71b93-530">エラー</span><span class="sxs-lookup"><span data-stu-id="71b93-530">Errors</span></span>

|<span data-ttu-id="71b93-531">エラー コード</span><span class="sxs-lookup"><span data-stu-id="71b93-531">Error code</span></span>|<span data-ttu-id="71b93-532">説明</span><span class="sxs-lookup"><span data-stu-id="71b93-532">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="71b93-533">要求が失敗しました。</span><span class="sxs-lookup"><span data-stu-id="71b93-533">The request has failed.</span></span> <span data-ttu-id="71b93-534">HTTP エラーコードの diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="71b93-534">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="71b93-535">Exchange サーバーがエラーを返しました。</span><span class="sxs-lookup"><span data-stu-id="71b93-535">The Exchange server returned an error.</span></span> <span data-ttu-id="71b93-536">詳細については、diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="71b93-536">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="71b93-537">ユーザーはネットワークに接続されていません。</span><span class="sxs-lookup"><span data-stu-id="71b93-537">The user is no longer connected to the network.</span></span> <span data-ttu-id="71b93-538">ネットワーク接続を確認し、やり直してください。</span><span class="sxs-lookup"><span data-stu-id="71b93-538">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="71b93-539">要件</span><span class="sxs-lookup"><span data-stu-id="71b93-539">Requirements</span></span>

|<span data-ttu-id="71b93-540">要件</span><span class="sxs-lookup"><span data-stu-id="71b93-540">Requirement</span></span>| <span data-ttu-id="71b93-541">値</span><span class="sxs-lookup"><span data-stu-id="71b93-541">Value</span></span>|
|---|---|
|[<span data-ttu-id="71b93-542">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="71b93-542">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="71b93-543">1.5</span><span class="sxs-lookup"><span data-stu-id="71b93-543">1.5</span></span> |
|[<span data-ttu-id="71b93-544">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="71b93-544">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="71b93-545">ReadItem</span><span class="sxs-lookup"><span data-stu-id="71b93-545">ReadItem</span></span>|
|[<span data-ttu-id="71b93-546">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="71b93-546">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="71b93-547">新規作成と閲覧</span><span class="sxs-lookup"><span data-stu-id="71b93-547">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="71b93-548">例</span><span class="sxs-lookup"><span data-stu-id="71b93-548">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="71b93-549">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="71b93-549">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="71b93-550">Exchange Server から添付ファイルやアイテムを取得するために使うトークンを含む文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="71b93-550">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="71b93-p140">`getCallbackTokenAsync` メソッドは、ユーザーのメールボックスをホストする Exchange Server から不透明なトークンを取得する非同期の呼び出しを行います。コールバック トークンの有効期間は 5 分です。</span><span class="sxs-lookup"><span data-stu-id="71b93-p140">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="71b93-553">トークンと、添付ファイル識別子またはアイテム識別子の両方をサードパーティのシステムに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="71b93-553">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="71b93-554">サードパーティシステムは、トークンをベアラー認証トークンとして使用して、Exchange Web サービス (EWS) の[Getattachment](/exchange/client-developer/web-service-reference/getattachment-operation)操作または[GetItem](/exchange/client-developer/web-service-reference/getitem-operation)操作を呼び出して、添付ファイルまたはアイテムを返します。</span><span class="sxs-lookup"><span data-stu-id="71b93-554">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item.</span></span> <span data-ttu-id="71b93-555">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="71b93-555">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="71b93-556">読み取りモード`getCallbackTokenAsync`でメソッドを呼び出すには、 **ReadItem**の最低限のアクセス許可レベルが必要です。</span><span class="sxs-lookup"><span data-stu-id="71b93-556">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="71b93-557">新規`getCallbackTokenAsync`作成モードで呼び出しを行うには、アイテムを保存しておく必要があります。</span><span class="sxs-lookup"><span data-stu-id="71b93-557">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="71b93-558">この[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback)メソッドには、 **readwriteitem**の最小アクセス許可レベルが必要です。</span><span class="sxs-lookup"><span data-stu-id="71b93-558">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

##### <a name="parameters"></a><span data-ttu-id="71b93-559">パラメーター</span><span class="sxs-lookup"><span data-stu-id="71b93-559">Parameters</span></span>

|<span data-ttu-id="71b93-560">名前</span><span class="sxs-lookup"><span data-stu-id="71b93-560">Name</span></span>| <span data-ttu-id="71b93-561">種類</span><span class="sxs-lookup"><span data-stu-id="71b93-561">Type</span></span>| <span data-ttu-id="71b93-562">属性</span><span class="sxs-lookup"><span data-stu-id="71b93-562">Attributes</span></span>| <span data-ttu-id="71b93-563">説明</span><span class="sxs-lookup"><span data-stu-id="71b93-563">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="71b93-564">関数</span><span class="sxs-lookup"><span data-stu-id="71b93-564">function</span></span>||<span data-ttu-id="71b93-565">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="71b93-565">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="71b93-566">トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="71b93-566">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="71b93-567">エラーが発生した場合、 `asyncResult.error` および `asyncResult.diagnostics` のプロパティで追加情報が提供される場合があります。</span><span class="sxs-lookup"><span data-stu-id="71b93-567">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="71b93-568">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="71b93-568">Object</span></span>| <span data-ttu-id="71b93-569">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="71b93-569">&lt;optional&gt;</span></span>|<span data-ttu-id="71b93-570">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="71b93-570">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="71b93-571">エラー</span><span class="sxs-lookup"><span data-stu-id="71b93-571">Errors</span></span>

|<span data-ttu-id="71b93-572">エラー コード</span><span class="sxs-lookup"><span data-stu-id="71b93-572">Error code</span></span>|<span data-ttu-id="71b93-573">説明</span><span class="sxs-lookup"><span data-stu-id="71b93-573">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="71b93-574">要求が失敗しました。</span><span class="sxs-lookup"><span data-stu-id="71b93-574">The request has failed.</span></span> <span data-ttu-id="71b93-575">HTTP エラーコードの diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="71b93-575">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="71b93-576">Exchange サーバーがエラーを返しました。</span><span class="sxs-lookup"><span data-stu-id="71b93-576">The Exchange server returned an error.</span></span> <span data-ttu-id="71b93-577">詳細については、diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="71b93-577">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="71b93-578">ユーザーはネットワークに接続されていません。</span><span class="sxs-lookup"><span data-stu-id="71b93-578">The user is no longer connected to the network.</span></span> <span data-ttu-id="71b93-579">ネットワーク接続を確認し、やり直してください。</span><span class="sxs-lookup"><span data-stu-id="71b93-579">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="71b93-580">要件</span><span class="sxs-lookup"><span data-stu-id="71b93-580">Requirements</span></span>

|<span data-ttu-id="71b93-581">要件</span><span class="sxs-lookup"><span data-stu-id="71b93-581">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="71b93-582">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="71b93-582">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="71b93-583">1.0</span><span class="sxs-lookup"><span data-stu-id="71b93-583">1.0</span></span> | <span data-ttu-id="71b93-584">1.3</span><span class="sxs-lookup"><span data-stu-id="71b93-584">1.3</span></span> |
|[<span data-ttu-id="71b93-585">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="71b93-585">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="71b93-586">ReadItem</span><span class="sxs-lookup"><span data-stu-id="71b93-586">ReadItem</span></span> | <span data-ttu-id="71b93-587">ReadItem</span><span class="sxs-lookup"><span data-stu-id="71b93-587">ReadItem</span></span> |
|[<span data-ttu-id="71b93-588">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="71b93-588">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="71b93-589">読み取り</span><span class="sxs-lookup"><span data-stu-id="71b93-589">Read</span></span> | <span data-ttu-id="71b93-590">作成</span><span class="sxs-lookup"><span data-stu-id="71b93-590">Compose</span></span> |

##### <a name="example"></a><span data-ttu-id="71b93-591">例</span><span class="sxs-lookup"><span data-stu-id="71b93-591">Example</span></span>

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

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="71b93-592">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="71b93-592">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="71b93-593">ユーザーと Office アドインを識別するトークンを取得します。</span><span class="sxs-lookup"><span data-stu-id="71b93-593">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="71b93-594">`getUserIdentityTokenAsync` メソッドは、[アドインとユーザーをサード パーティのシステムで識別して認証](/outlook/add-ins/authentication)することのできるトークンを返します。</span><span class="sxs-lookup"><span data-stu-id="71b93-594">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="71b93-595">パラメーター</span><span class="sxs-lookup"><span data-stu-id="71b93-595">Parameters</span></span>

|<span data-ttu-id="71b93-596">名前</span><span class="sxs-lookup"><span data-stu-id="71b93-596">Name</span></span>| <span data-ttu-id="71b93-597">種類</span><span class="sxs-lookup"><span data-stu-id="71b93-597">Type</span></span>| <span data-ttu-id="71b93-598">属性</span><span class="sxs-lookup"><span data-stu-id="71b93-598">Attributes</span></span>| <span data-ttu-id="71b93-599">説明</span><span class="sxs-lookup"><span data-stu-id="71b93-599">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="71b93-600">関数</span><span class="sxs-lookup"><span data-stu-id="71b93-600">function</span></span>||<span data-ttu-id="71b93-601">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="71b93-601">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="71b93-602">トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="71b93-602">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="71b93-603">エラーが発生した場合、 `asyncResult.error` および `asyncResult.diagnostics` のプロパティで追加情報が提供される場合があります。</span><span class="sxs-lookup"><span data-stu-id="71b93-603">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="71b93-604">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="71b93-604">Object</span></span>| <span data-ttu-id="71b93-605">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="71b93-605">&lt;optional&gt;</span></span>|<span data-ttu-id="71b93-606">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="71b93-606">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="71b93-607">エラー</span><span class="sxs-lookup"><span data-stu-id="71b93-607">Errors</span></span>

|<span data-ttu-id="71b93-608">エラー コード</span><span class="sxs-lookup"><span data-stu-id="71b93-608">Error code</span></span>|<span data-ttu-id="71b93-609">説明</span><span class="sxs-lookup"><span data-stu-id="71b93-609">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="71b93-610">要求が失敗しました。</span><span class="sxs-lookup"><span data-stu-id="71b93-610">The request has failed.</span></span> <span data-ttu-id="71b93-611">HTTP エラーコードの diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="71b93-611">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="71b93-612">Exchange サーバーがエラーを返しました。</span><span class="sxs-lookup"><span data-stu-id="71b93-612">The Exchange server returned an error.</span></span> <span data-ttu-id="71b93-613">詳細については、diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="71b93-613">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="71b93-614">ユーザーはネットワークに接続されていません。</span><span class="sxs-lookup"><span data-stu-id="71b93-614">The user is no longer connected to the network.</span></span> <span data-ttu-id="71b93-615">ネットワーク接続を確認し、やり直してください。</span><span class="sxs-lookup"><span data-stu-id="71b93-615">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="71b93-616">要件</span><span class="sxs-lookup"><span data-stu-id="71b93-616">Requirements</span></span>

|<span data-ttu-id="71b93-617">要件</span><span class="sxs-lookup"><span data-stu-id="71b93-617">Requirement</span></span>| <span data-ttu-id="71b93-618">値</span><span class="sxs-lookup"><span data-stu-id="71b93-618">Value</span></span>|
|---|---|
|[<span data-ttu-id="71b93-619">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="71b93-619">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="71b93-620">1.0</span><span class="sxs-lookup"><span data-stu-id="71b93-620">1.0</span></span>|
|[<span data-ttu-id="71b93-621">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="71b93-621">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="71b93-622">ReadItem</span><span class="sxs-lookup"><span data-stu-id="71b93-622">ReadItem</span></span>|
|[<span data-ttu-id="71b93-623">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="71b93-623">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="71b93-624">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="71b93-624">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="71b93-625">例</span><span class="sxs-lookup"><span data-stu-id="71b93-625">Example</span></span>

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

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="71b93-626">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="71b93-626">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="71b93-627">ユーザーのメールボックスをホストしている Exchange サーバー上の Exchange Web サービス (EWS) のサービスに対して非同期の要求を行います。</span><span class="sxs-lookup"><span data-stu-id="71b93-627">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="71b93-628">このメソッドは、次のシナリオではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="71b93-628">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="71b93-629">Outlook on iOS または Android の場合</span><span class="sxs-lookup"><span data-stu-id="71b93-629">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="71b93-630">アドインが Gmail のメールボックスに読み込まれる場合</span><span class="sxs-lookup"><span data-stu-id="71b93-630">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="71b93-631">このような場合は、アドインでは [REST API を使用](/outlook/add-ins/use-rest-api)して、代わりにユーザーのメールボックスにアクセスする必要があります。</span><span class="sxs-lookup"><span data-stu-id="71b93-631">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="71b93-p149">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange. See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span><span class="sxs-lookup"><span data-stu-id="71b93-p149">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange. See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="71b93-634">`makeEwsRequestAsync` メソッドでは、フォルダー関連アイテムを要求できません。</span><span class="sxs-lookup"><span data-stu-id="71b93-634">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="71b93-635">XML 要求では UTF-8 エンコードを指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="71b93-635">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="71b93-p150">`makeEwsRequestAsync` メソッドを使用するには、アドインに **ReadWriteMailbox** アクセス許可が必要です。**ReadWriteMailbox** アクセス許可と、`makeEwsRequestAsync` メソッドで呼び出せる EWS 操作の使い方については、「[ユーザーのメールボックスへのメール アドイン アクセスのアクセス許可を指定する](/outlook/add-ins/understanding-outlook-add-in-permissions)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="71b93-p150">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="71b93-638">サーバー管理者は、クライアント アクセス サーバーの EWS ディレクトリで `OAuthAuthentication` を true に設定して、`makeEwsRequestAsync` メソッドで EWS 要求を行うことができるようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="71b93-638">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="71b93-639">バージョンの相違点</span><span class="sxs-lookup"><span data-stu-id="71b93-639">Version differences</span></span>

<span data-ttu-id="71b93-640">バージョン 15.0.4535.1004 より前のバージョンの Outlook で実行しているメール アプリで `makeEwsRequestAsync` メソッドを使用する場合には、エンコード値を `ISO-8859-1` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="71b93-640">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="71b93-641">Outlook on the web でメール アプリを実行している場合は、エンコード値を設定する必要はありません。</span><span class="sxs-lookup"><span data-stu-id="71b93-641">You do not need to set the encoding value when your mail app is running in Outlook on the web.</span></span> <span data-ttu-id="71b93-642">メールアプリが web 上の Outlook またはデスクトップクライアントで実行されているかどうかは、mailbox プロパティを使用して判断できます。</span><span class="sxs-lookup"><span data-stu-id="71b93-642">You can determine whether your mail app is running in Outlook on the web or a desktop client by using the mailbox.diagnostics.hostName property.</span></span> <span data-ttu-id="71b93-643">mailbox.diagnostics.hostVersion プロパティを使って、どのバージョンの Outlook を使って実行しているかを確認できます。</span><span class="sxs-lookup"><span data-stu-id="71b93-643">You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="71b93-644">パラメーター</span><span class="sxs-lookup"><span data-stu-id="71b93-644">Parameters</span></span>

|<span data-ttu-id="71b93-645">名前</span><span class="sxs-lookup"><span data-stu-id="71b93-645">Name</span></span>| <span data-ttu-id="71b93-646">種類</span><span class="sxs-lookup"><span data-stu-id="71b93-646">Type</span></span>| <span data-ttu-id="71b93-647">属性</span><span class="sxs-lookup"><span data-stu-id="71b93-647">Attributes</span></span>| <span data-ttu-id="71b93-648">説明</span><span class="sxs-lookup"><span data-stu-id="71b93-648">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="71b93-649">String</span><span class="sxs-lookup"><span data-stu-id="71b93-649">String</span></span>||<span data-ttu-id="71b93-650">EWS 要求です。</span><span class="sxs-lookup"><span data-stu-id="71b93-650">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="71b93-651">function</span><span class="sxs-lookup"><span data-stu-id="71b93-651">function</span></span>||<span data-ttu-id="71b93-652">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="71b93-652">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="71b93-p152">The XML result of the EWS call is provided as a string in the `asyncResult.value` property. If the result exceeds 1 MB in size, an error message is returned instead.</span><span class="sxs-lookup"><span data-stu-id="71b93-p152">The XML result of the EWS call is provided as a string in the `asyncResult.value` property. If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="71b93-655">Object</span><span class="sxs-lookup"><span data-stu-id="71b93-655">Object</span></span>| <span data-ttu-id="71b93-656">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="71b93-656">&lt;optional&gt;</span></span>|<span data-ttu-id="71b93-657">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="71b93-657">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="71b93-658">要件</span><span class="sxs-lookup"><span data-stu-id="71b93-658">Requirements</span></span>

|<span data-ttu-id="71b93-659">要件</span><span class="sxs-lookup"><span data-stu-id="71b93-659">Requirement</span></span>| <span data-ttu-id="71b93-660">値</span><span class="sxs-lookup"><span data-stu-id="71b93-660">Value</span></span>|
|---|---|
|[<span data-ttu-id="71b93-661">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="71b93-661">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="71b93-662">1.0</span><span class="sxs-lookup"><span data-stu-id="71b93-662">1.0</span></span>|
|[<span data-ttu-id="71b93-663">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="71b93-663">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="71b93-664">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="71b93-664">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="71b93-665">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="71b93-665">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="71b93-666">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="71b93-666">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="71b93-667">例</span><span class="sxs-lookup"><span data-stu-id="71b93-667">Example</span></span>

<span data-ttu-id="71b93-668">次の例は、`GetItem` 操作を使ってアイテムの件名を取得するため、`makeEwsRequestAsync` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="71b93-668">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="71b93-669">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="71b93-669">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="71b93-670">サポートされているイベントの種類のイベント ハンドラーを削除します。</span><span class="sxs-lookup"><span data-stu-id="71b93-670">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="71b93-671">現在、サポートされている`Office.EventType.ItemChanged`イベント`Office.EventType.OfficeThemeChanged`の種類はおよびです。</span><span class="sxs-lookup"><span data-stu-id="71b93-671">Currently, the supported event types are `Office.EventType.ItemChanged` and `Office.EventType.OfficeThemeChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="71b93-672">パラメーター</span><span class="sxs-lookup"><span data-stu-id="71b93-672">Parameters</span></span>

| <span data-ttu-id="71b93-673">名前</span><span class="sxs-lookup"><span data-stu-id="71b93-673">Name</span></span> | <span data-ttu-id="71b93-674">種類</span><span class="sxs-lookup"><span data-stu-id="71b93-674">Type</span></span> | <span data-ttu-id="71b93-675">属性</span><span class="sxs-lookup"><span data-stu-id="71b93-675">Attributes</span></span> | <span data-ttu-id="71b93-676">説明</span><span class="sxs-lookup"><span data-stu-id="71b93-676">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="71b93-677">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="71b93-677">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="71b93-678">ハンドラーを取り消すイベント。</span><span class="sxs-lookup"><span data-stu-id="71b93-678">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="71b93-679">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="71b93-679">Object</span></span> | <span data-ttu-id="71b93-680">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="71b93-680">&lt;optional&gt;</span></span> | <span data-ttu-id="71b93-681">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="71b93-681">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="71b93-682">Object</span><span class="sxs-lookup"><span data-stu-id="71b93-682">Object</span></span> | <span data-ttu-id="71b93-683">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="71b93-683">&lt;optional&gt;</span></span> | <span data-ttu-id="71b93-684">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="71b93-684">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="71b93-685">function</span><span class="sxs-lookup"><span data-stu-id="71b93-685">function</span></span>| <span data-ttu-id="71b93-686">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="71b93-686">&lt;optional&gt;</span></span>|<span data-ttu-id="71b93-687">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="71b93-687">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="71b93-688">要件</span><span class="sxs-lookup"><span data-stu-id="71b93-688">Requirements</span></span>

|<span data-ttu-id="71b93-689">要件</span><span class="sxs-lookup"><span data-stu-id="71b93-689">Requirement</span></span>| <span data-ttu-id="71b93-690">値</span><span class="sxs-lookup"><span data-stu-id="71b93-690">Value</span></span>|
|---|---|
|[<span data-ttu-id="71b93-691">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="71b93-691">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="71b93-692">1.5</span><span class="sxs-lookup"><span data-stu-id="71b93-692">1.5</span></span> |
|[<span data-ttu-id="71b93-693">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="71b93-693">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="71b93-694">ReadItem</span><span class="sxs-lookup"><span data-stu-id="71b93-694">ReadItem</span></span> |
|[<span data-ttu-id="71b93-695">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="71b93-695">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="71b93-696">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="71b93-696">Compose or Read</span></span>|
