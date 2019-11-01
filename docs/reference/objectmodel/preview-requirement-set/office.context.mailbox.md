---
title: Office のメールボックス-プレビュー要件セット
description: ''
ms.date: 10/30/2019
localization_priority: Normal
ms.openlocfilehash: ff649029713984b32e817bbeaf7c59a48cc5b023
ms.sourcegitcommit: e989096f3d19761bf8477c585cde20b3f8e0b90d
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/31/2019
ms.locfileid: "37902110"
---
# <a name="mailbox"></a><span data-ttu-id="cff5c-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="cff5c-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="cff5c-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="cff5c-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="cff5c-104">Microsoft Outlook の Outlook アドイン オブジェクト モデルへのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="cff5c-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="cff5c-105">要件</span><span class="sxs-lookup"><span data-stu-id="cff5c-105">Requirements</span></span>

|<span data-ttu-id="cff5c-106">要件</span><span class="sxs-lookup"><span data-stu-id="cff5c-106">Requirement</span></span>| <span data-ttu-id="cff5c-107">値</span><span class="sxs-lookup"><span data-stu-id="cff5c-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="cff5c-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cff5c-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cff5c-109">1.0</span><span class="sxs-lookup"><span data-stu-id="cff5c-109">1.0</span></span>|
|[<span data-ttu-id="cff5c-110">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cff5c-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cff5c-111">制限あり</span><span class="sxs-lookup"><span data-stu-id="cff5c-111">Restricted</span></span>|
|[<span data-ttu-id="cff5c-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cff5c-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cff5c-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="cff5c-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="cff5c-114">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="cff5c-114">Members and methods</span></span>

| <span data-ttu-id="cff5c-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="cff5c-115">Member</span></span> | <span data-ttu-id="cff5c-116">種類</span><span class="sxs-lookup"><span data-stu-id="cff5c-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="cff5c-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="cff5c-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="cff5c-118">メンバー</span><span class="sxs-lookup"><span data-stu-id="cff5c-118">Member</span></span> |
| [<span data-ttu-id="cff5c-119">masterCategories</span><span class="sxs-lookup"><span data-stu-id="cff5c-119">masterCategories</span></span>](#mastercategories-mastercategories) | <span data-ttu-id="cff5c-120">メンバー</span><span class="sxs-lookup"><span data-stu-id="cff5c-120">Member</span></span> |
| [<span data-ttu-id="cff5c-121">restUrl</span><span class="sxs-lookup"><span data-stu-id="cff5c-121">restUrl</span></span>](#resturl-string) | <span data-ttu-id="cff5c-122">メンバー</span><span class="sxs-lookup"><span data-stu-id="cff5c-122">Member</span></span> |
| [<span data-ttu-id="cff5c-123">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="cff5c-123">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="cff5c-124">メソッド</span><span class="sxs-lookup"><span data-stu-id="cff5c-124">Method</span></span> |
| [<span data-ttu-id="cff5c-125">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="cff5c-125">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="cff5c-126">メソッド</span><span class="sxs-lookup"><span data-stu-id="cff5c-126">Method</span></span> |
| [<span data-ttu-id="cff5c-127">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="cff5c-127">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="cff5c-128">メソッド</span><span class="sxs-lookup"><span data-stu-id="cff5c-128">Method</span></span> |
| [<span data-ttu-id="cff5c-129">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="cff5c-129">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="cff5c-130">メソッド</span><span class="sxs-lookup"><span data-stu-id="cff5c-130">Method</span></span> |
| [<span data-ttu-id="cff5c-131">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="cff5c-131">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="cff5c-132">メソッド</span><span class="sxs-lookup"><span data-stu-id="cff5c-132">Method</span></span> |
| [<span data-ttu-id="cff5c-133">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="cff5c-133">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="cff5c-134">メソッド</span><span class="sxs-lookup"><span data-stu-id="cff5c-134">Method</span></span> |
| [<span data-ttu-id="cff5c-135">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="cff5c-135">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="cff5c-136">メソッド</span><span class="sxs-lookup"><span data-stu-id="cff5c-136">Method</span></span> |
| [<span data-ttu-id="cff5c-137">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="cff5c-137">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="cff5c-138">メソッド</span><span class="sxs-lookup"><span data-stu-id="cff5c-138">Method</span></span> |
| [<span data-ttu-id="cff5c-139">displayNewMessageForm</span><span class="sxs-lookup"><span data-stu-id="cff5c-139">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="cff5c-140">メソッド</span><span class="sxs-lookup"><span data-stu-id="cff5c-140">Method</span></span> |
| [<span data-ttu-id="cff5c-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="cff5c-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="cff5c-142">メソッド</span><span class="sxs-lookup"><span data-stu-id="cff5c-142">Method</span></span> |
| [<span data-ttu-id="cff5c-143">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="cff5c-143">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="cff5c-144">メソッド</span><span class="sxs-lookup"><span data-stu-id="cff5c-144">Method</span></span> |
| [<span data-ttu-id="cff5c-145">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="cff5c-145">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="cff5c-146">メソッド</span><span class="sxs-lookup"><span data-stu-id="cff5c-146">Method</span></span> |
| [<span data-ttu-id="cff5c-147">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="cff5c-147">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="cff5c-148">メソッド</span><span class="sxs-lookup"><span data-stu-id="cff5c-148">Method</span></span> |
| [<span data-ttu-id="cff5c-149">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="cff5c-149">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="cff5c-150">メソッド</span><span class="sxs-lookup"><span data-stu-id="cff5c-150">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="cff5c-151">名前空間</span><span class="sxs-lookup"><span data-stu-id="cff5c-151">Namespaces</span></span>

<span data-ttu-id="cff5c-152">[diagnostics](Office.context.mailbox.diagnostics.md):Outlook アドインに診断情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="cff5c-152">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="cff5c-153">[item](Office.context.mailbox.item.md):Outlook アドインのメッセージや予定にアクセスするためのメソッドとプロパティを提供します。</span><span class="sxs-lookup"><span data-stu-id="cff5c-153">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="cff5c-154">[userProfile](Office.context.mailbox.userProfile.md):Outlook アドインのユーザーに関する情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="cff5c-154">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="cff5c-155">Members</span><span class="sxs-lookup"><span data-stu-id="cff5c-155">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="cff5c-156">ewsUrl: String</span><span class="sxs-lookup"><span data-stu-id="cff5c-156">ewsUrl: String</span></span>

<span data-ttu-id="cff5c-p101">このメール アカウントの Exchange Web サービス (EWS) エンドポイントの URL を取得します。読み取りモードのみです。</span><span class="sxs-lookup"><span data-stu-id="cff5c-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="cff5c-159">このメンバーは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="cff5c-159">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="cff5c-p102">
  `ewsUrl\` 値は、リモート サービスで、ユーザーのメールボックスに EWS 呼び出しを行うために使うことができます。たとえば、[選択したアイテムから添付ファイルを取得する](/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="cff5c-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="cff5c-162">アプリが閲覧モードで `ewsUrl` メンバーを呼び出すには、アプリのマニフェスト内に **ReadItem** アクセス許可が指定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="cff5c-162">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="cff5c-p103">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出してから、`ewsUrl` メンバーを使用する必要があります。アプリには、`saveAsync` メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="cff5c-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="cff5c-165">型</span><span class="sxs-lookup"><span data-stu-id="cff5c-165">Type</span></span>

*   <span data-ttu-id="cff5c-166">String</span><span class="sxs-lookup"><span data-stu-id="cff5c-166">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="cff5c-167">要件</span><span class="sxs-lookup"><span data-stu-id="cff5c-167">Requirements</span></span>

|<span data-ttu-id="cff5c-168">要件</span><span class="sxs-lookup"><span data-stu-id="cff5c-168">Requirement</span></span>| <span data-ttu-id="cff5c-169">値</span><span class="sxs-lookup"><span data-stu-id="cff5c-169">Value</span></span>|
|---|---|
|[<span data-ttu-id="cff5c-170">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cff5c-170">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cff5c-171">1.0</span><span class="sxs-lookup"><span data-stu-id="cff5c-171">1.0</span></span>|
|[<span data-ttu-id="cff5c-172">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cff5c-172">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cff5c-173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cff5c-173">ReadItem</span></span>|
|[<span data-ttu-id="cff5c-174">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cff5c-174">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cff5c-175">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="cff5c-175">Compose or Read</span></span>|

<br>

---
---

#### <a name="mastercategories-mastercategoriesjavascriptapioutlookofficemastercategories"></a><span data-ttu-id="cff5c-176">masterCategories: [Mastercategories](/javascript/api/outlook/office.mastercategories)</span><span class="sxs-lookup"><span data-stu-id="cff5c-176">masterCategories: [MasterCategories](/javascript/api/outlook/office.mastercategories)</span></span>

<span data-ttu-id="cff5c-177">このメールボックスのカテゴリマスターリストを管理するためのメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="cff5c-177">Gets an object that provides methods to manage the categories master list on this mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="cff5c-178">このメンバーは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="cff5c-178">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="cff5c-179">種類</span><span class="sxs-lookup"><span data-stu-id="cff5c-179">Type</span></span>

*   [<span data-ttu-id="cff5c-180">MasterCategories</span><span class="sxs-lookup"><span data-stu-id="cff5c-180">MasterCategories</span></span>](/javascript/api/outlook/office.mastercategories)

##### <a name="requirements"></a><span data-ttu-id="cff5c-181">要件</span><span class="sxs-lookup"><span data-stu-id="cff5c-181">Requirements</span></span>

|<span data-ttu-id="cff5c-182">要件</span><span class="sxs-lookup"><span data-stu-id="cff5c-182">Requirement</span></span>| <span data-ttu-id="cff5c-183">値</span><span class="sxs-lookup"><span data-stu-id="cff5c-183">Value</span></span>|
|---|---|
|[<span data-ttu-id="cff5c-184">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cff5c-184">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cff5c-185">1.8</span><span class="sxs-lookup"><span data-stu-id="cff5c-185">1.8</span></span> |
|[<span data-ttu-id="cff5c-186">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cff5c-186">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cff5c-187">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="cff5c-187">ReadWriteMailbox</span></span> |
|[<span data-ttu-id="cff5c-188">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cff5c-188">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cff5c-189">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="cff5c-189">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="cff5c-190">例</span><span class="sxs-lookup"><span data-stu-id="cff5c-190">Example</span></span>

<span data-ttu-id="cff5c-191">この例では、このメールボックスのカテゴリマスターリストを取得します。</span><span class="sxs-lookup"><span data-stu-id="cff5c-191">This example gets the categories master list for this mailbox.</span></span>

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

#### <a name="resturl-string"></a><span data-ttu-id="cff5c-192">restUrl: String</span><span class="sxs-lookup"><span data-stu-id="cff5c-192">restUrl: String</span></span>

<span data-ttu-id="cff5c-193">この電子メール アカウントの REST エンドポイントの URL を取得します。</span><span class="sxs-lookup"><span data-stu-id="cff5c-193">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="cff5c-194">`restUrl` 値は、ユーザーのメールボックスに [REST API](/outlook/rest/) 呼び出しを行うために使用できます。</span><span class="sxs-lookup"><span data-stu-id="cff5c-194">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="cff5c-195">アプリが閲覧モードで `restUrl` メンバーを呼び出すには、アプリのマニフェスト内に **ReadItem** アクセス許可が指定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="cff5c-195">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="cff5c-p104">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出してから、`restUrl` メンバーを使用する必要があります。アプリには、`saveAsync` メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="cff5c-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="cff5c-198">型</span><span class="sxs-lookup"><span data-stu-id="cff5c-198">Type</span></span>

*   <span data-ttu-id="cff5c-199">String</span><span class="sxs-lookup"><span data-stu-id="cff5c-199">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="cff5c-200">要件</span><span class="sxs-lookup"><span data-stu-id="cff5c-200">Requirements</span></span>

|<span data-ttu-id="cff5c-201">要件</span><span class="sxs-lookup"><span data-stu-id="cff5c-201">Requirement</span></span>| <span data-ttu-id="cff5c-202">値</span><span class="sxs-lookup"><span data-stu-id="cff5c-202">Value</span></span>|
|---|---|
|[<span data-ttu-id="cff5c-203">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cff5c-203">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cff5c-204">1.5</span><span class="sxs-lookup"><span data-stu-id="cff5c-204">1.5</span></span> |
|[<span data-ttu-id="cff5c-205">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cff5c-205">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cff5c-206">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cff5c-206">ReadItem</span></span>|
|[<span data-ttu-id="cff5c-207">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cff5c-207">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cff5c-208">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="cff5c-208">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="cff5c-209">メソッド</span><span class="sxs-lookup"><span data-stu-id="cff5c-209">Methods</span></span>

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="cff5c-210">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="cff5c-210">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="cff5c-211">サポートされているイベントのイベント ハンドラーを追加します。</span><span class="sxs-lookup"><span data-stu-id="cff5c-211">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="cff5c-212">現在、サポートされている`Office.EventType.ItemChanged`イベント`Office.EventType.OfficeThemeChanged`の種類はおよびです。</span><span class="sxs-lookup"><span data-stu-id="cff5c-212">Currently, the supported event types are `Office.EventType.ItemChanged` and `Office.EventType.OfficeThemeChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cff5c-213">パラメーター</span><span class="sxs-lookup"><span data-stu-id="cff5c-213">Parameters</span></span>

| <span data-ttu-id="cff5c-214">名前</span><span class="sxs-lookup"><span data-stu-id="cff5c-214">Name</span></span> | <span data-ttu-id="cff5c-215">種類</span><span class="sxs-lookup"><span data-stu-id="cff5c-215">Type</span></span> | <span data-ttu-id="cff5c-216">属性</span><span class="sxs-lookup"><span data-stu-id="cff5c-216">Attributes</span></span> | <span data-ttu-id="cff5c-217">説明</span><span class="sxs-lookup"><span data-stu-id="cff5c-217">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="cff5c-218">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="cff5c-218">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="cff5c-219">ハンドラーを呼び出す必要のあるイベント。</span><span class="sxs-lookup"><span data-stu-id="cff5c-219">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="cff5c-220">Function</span><span class="sxs-lookup"><span data-stu-id="cff5c-220">Function</span></span> || <span data-ttu-id="cff5c-p105">イベントを処理する関数。関数は、オブジェクト リテラルである単一パラメーターを受け入れる必要があります。パラメーターの `type` プロパティは、`addHandlerAsync` に渡される `eventType` パラメーターと一致します。</span><span class="sxs-lookup"><span data-stu-id="cff5c-p105">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="cff5c-224">Object</span><span class="sxs-lookup"><span data-stu-id="cff5c-224">Object</span></span> | <span data-ttu-id="cff5c-225">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="cff5c-225">&lt;optional&gt;</span></span> | <span data-ttu-id="cff5c-226">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="cff5c-226">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="cff5c-227">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="cff5c-227">Object</span></span> | <span data-ttu-id="cff5c-228">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="cff5c-228">&lt;optional&gt;</span></span> | <span data-ttu-id="cff5c-229">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="cff5c-229">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="cff5c-230">function</span><span class="sxs-lookup"><span data-stu-id="cff5c-230">function</span></span>| <span data-ttu-id="cff5c-231">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="cff5c-231">&lt;optional&gt;</span></span>|<span data-ttu-id="cff5c-232">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="cff5c-232">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cff5c-233">要件</span><span class="sxs-lookup"><span data-stu-id="cff5c-233">Requirements</span></span>

|<span data-ttu-id="cff5c-234">要件</span><span class="sxs-lookup"><span data-stu-id="cff5c-234">Requirement</span></span>| <span data-ttu-id="cff5c-235">値</span><span class="sxs-lookup"><span data-stu-id="cff5c-235">Value</span></span>|
|---|---|
|[<span data-ttu-id="cff5c-236">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cff5c-236">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cff5c-237">1.5</span><span class="sxs-lookup"><span data-stu-id="cff5c-237">1.5</span></span> |
|[<span data-ttu-id="cff5c-238">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cff5c-238">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cff5c-239">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cff5c-239">ReadItem</span></span> |
|[<span data-ttu-id="cff5c-240">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cff5c-240">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cff5c-241">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="cff5c-241">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cff5c-242">例</span><span class="sxs-lookup"><span data-stu-id="cff5c-242">Example</span></span>

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

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="cff5c-243">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="cff5c-243">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="cff5c-244">REST 形式のアイテム ID を EWS 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="cff5c-244">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="cff5c-245">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="cff5c-245">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="cff5c-p106">REST API ([Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) や [Microsoft Graph](https://graph.microsoft.io/) など) で取得されたアイテム ID は、Exchange Web サービス (EWS) に使用される形式とは異なる形式を使用します。`convertToEwsId` メソッドは、REST 形式の ID を EWS 用の適切な形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="cff5c-p106">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cff5c-248">パラメーター</span><span class="sxs-lookup"><span data-stu-id="cff5c-248">Parameters</span></span>

|<span data-ttu-id="cff5c-249">名前</span><span class="sxs-lookup"><span data-stu-id="cff5c-249">Name</span></span>| <span data-ttu-id="cff5c-250">型</span><span class="sxs-lookup"><span data-stu-id="cff5c-250">Type</span></span>| <span data-ttu-id="cff5c-251">説明</span><span class="sxs-lookup"><span data-stu-id="cff5c-251">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="cff5c-252">String</span><span class="sxs-lookup"><span data-stu-id="cff5c-252">String</span></span>|<span data-ttu-id="cff5c-253">Outlook REST API 形式のアイテム ID</span><span class="sxs-lookup"><span data-stu-id="cff5c-253">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="cff5c-254">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="cff5c-254">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion)|<span data-ttu-id="cff5c-255">アイテム ID の取得に使用された Outlook REST API のバージョンを示す値。</span><span class="sxs-lookup"><span data-stu-id="cff5c-255">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cff5c-256">要件</span><span class="sxs-lookup"><span data-stu-id="cff5c-256">Requirements</span></span>

|<span data-ttu-id="cff5c-257">要件</span><span class="sxs-lookup"><span data-stu-id="cff5c-257">Requirement</span></span>| <span data-ttu-id="cff5c-258">値</span><span class="sxs-lookup"><span data-stu-id="cff5c-258">Value</span></span>|
|---|---|
|[<span data-ttu-id="cff5c-259">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cff5c-259">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cff5c-260">1.3</span><span class="sxs-lookup"><span data-stu-id="cff5c-260">1.3</span></span>|
|[<span data-ttu-id="cff5c-261">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cff5c-261">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cff5c-262">制限あり</span><span class="sxs-lookup"><span data-stu-id="cff5c-262">Restricted</span></span>|
|[<span data-ttu-id="cff5c-263">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cff5c-263">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cff5c-264">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="cff5c-264">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="cff5c-265">戻り値:</span><span class="sxs-lookup"><span data-stu-id="cff5c-265">Returns:</span></span>

<span data-ttu-id="cff5c-266">型:String</span><span class="sxs-lookup"><span data-stu-id="cff5c-266">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="cff5c-267">例</span><span class="sxs-lookup"><span data-stu-id="cff5c-267">Example</span></span>

```js
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime"></a><span data-ttu-id="cff5c-268">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="cff5c-268">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime)}</span></span>

<span data-ttu-id="cff5c-269">クライアントのローカル時間で時間情報が含まれている辞書を取得します。</span><span class="sxs-lookup"><span data-stu-id="cff5c-269">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="cff5c-p107">Outlook on the web または Outlook デスクトップのメールアプリでは、日付と時刻に異なるタイムゾーンを使用できます。Outlook デスクトップは、クライアント コンピューターのタイム ゾーンを使用します。Outlook on the web は、Exchange 管理センター (EAC) で設定されたタイム ゾーンを使用します。日付と時刻の値は、ユーザー インターフェイスに表示される値が、常にユーザーが期待するタイム ゾーンと一致するように処理する必要があります。</span><span class="sxs-lookup"><span data-stu-id="cff5c-p107">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times. Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="cff5c-p108">Outlook デスクトップ クライアントでメール アプリを実行している場合、`convertToLocalClientTime` メソッドは、クライアント コンピューターのタイム ゾーンに設定された値のディクショナリ オブジェクトを返します。Outlook on the web でメール アプリを実行している場合、`convertToLocalClientTime` メソッドは、EAC で指定したタイム ゾーンに設定された値のディクショナリ オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="cff5c-p108">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cff5c-275">パラメーター</span><span class="sxs-lookup"><span data-stu-id="cff5c-275">Parameters</span></span>

|<span data-ttu-id="cff5c-276">名前</span><span class="sxs-lookup"><span data-stu-id="cff5c-276">Name</span></span>| <span data-ttu-id="cff5c-277">型</span><span class="sxs-lookup"><span data-stu-id="cff5c-277">Type</span></span>| <span data-ttu-id="cff5c-278">説明</span><span class="sxs-lookup"><span data-stu-id="cff5c-278">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="cff5c-279">日付</span><span class="sxs-lookup"><span data-stu-id="cff5c-279">Date</span></span>|<span data-ttu-id="cff5c-280">日付オブジェクト</span><span class="sxs-lookup"><span data-stu-id="cff5c-280">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cff5c-281">要件</span><span class="sxs-lookup"><span data-stu-id="cff5c-281">Requirements</span></span>

|<span data-ttu-id="cff5c-282">要件</span><span class="sxs-lookup"><span data-stu-id="cff5c-282">Requirement</span></span>| <span data-ttu-id="cff5c-283">値</span><span class="sxs-lookup"><span data-stu-id="cff5c-283">Value</span></span>|
|---|---|
|[<span data-ttu-id="cff5c-284">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cff5c-284">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cff5c-285">1.0</span><span class="sxs-lookup"><span data-stu-id="cff5c-285">1.0</span></span>|
|[<span data-ttu-id="cff5c-286">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cff5c-286">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cff5c-287">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cff5c-287">ReadItem</span></span>|
|[<span data-ttu-id="cff5c-288">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cff5c-288">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cff5c-289">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="cff5c-289">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="cff5c-290">戻り値:</span><span class="sxs-lookup"><span data-stu-id="cff5c-290">Returns:</span></span>

<span data-ttu-id="cff5c-291">型:[LocalClientTime](/javascript/api/outlook/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="cff5c-291">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime)</span></span>

<br>

---
---

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="cff5c-292">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="cff5c-292">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="cff5c-293">EWS 形式のアイテム ID を REST 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="cff5c-293">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="cff5c-294">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="cff5c-294">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="cff5c-p109">EWS または `itemId` プロパティで取得されるアイテム ID は、REST API ([Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) や [Microsoft Graph](https://graph.microsoft.io/) など) に使用される形式とは異なる形式を使用します。`convertToRestId` メソッドは、EWS 形式の ID を REST 用の適切な形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="cff5c-p109">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cff5c-297">パラメーター</span><span class="sxs-lookup"><span data-stu-id="cff5c-297">Parameters</span></span>

|<span data-ttu-id="cff5c-298">名前</span><span class="sxs-lookup"><span data-stu-id="cff5c-298">Name</span></span>| <span data-ttu-id="cff5c-299">型</span><span class="sxs-lookup"><span data-stu-id="cff5c-299">Type</span></span>| <span data-ttu-id="cff5c-300">説明</span><span class="sxs-lookup"><span data-stu-id="cff5c-300">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="cff5c-301">String</span><span class="sxs-lookup"><span data-stu-id="cff5c-301">String</span></span>|<span data-ttu-id="cff5c-302">Exchange Web サービス (EWS) 形式のアイテム ID</span><span class="sxs-lookup"><span data-stu-id="cff5c-302">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="cff5c-303">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="cff5c-303">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion)|<span data-ttu-id="cff5c-304">変換後の ID を使用する Outlook REST API のバージョンを示す値。</span><span class="sxs-lookup"><span data-stu-id="cff5c-304">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cff5c-305">要件</span><span class="sxs-lookup"><span data-stu-id="cff5c-305">Requirements</span></span>

|<span data-ttu-id="cff5c-306">要件</span><span class="sxs-lookup"><span data-stu-id="cff5c-306">Requirement</span></span>| <span data-ttu-id="cff5c-307">値</span><span class="sxs-lookup"><span data-stu-id="cff5c-307">Value</span></span>|
|---|---|
|[<span data-ttu-id="cff5c-308">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cff5c-308">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cff5c-309">1.3</span><span class="sxs-lookup"><span data-stu-id="cff5c-309">1.3</span></span>|
|[<span data-ttu-id="cff5c-310">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cff5c-310">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cff5c-311">制限あり</span><span class="sxs-lookup"><span data-stu-id="cff5c-311">Restricted</span></span>|
|[<span data-ttu-id="cff5c-312">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cff5c-312">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cff5c-313">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="cff5c-313">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="cff5c-314">戻り値:</span><span class="sxs-lookup"><span data-stu-id="cff5c-314">Returns:</span></span>

<span data-ttu-id="cff5c-315">型:String</span><span class="sxs-lookup"><span data-stu-id="cff5c-315">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="cff5c-316">例</span><span class="sxs-lookup"><span data-stu-id="cff5c-316">Example</span></span>

```js
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="cff5c-317">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="cff5c-317">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="cff5c-318">時間情報が含まれているディクショナリから日付オブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="cff5c-318">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="cff5c-319">`convertToUtcClientTime` メソッドは、ローカルの日付と時刻を含むディクショナリを、ローカルの日付と時刻の正しい値を持つ日付オブジェクトに変換します。</span><span class="sxs-lookup"><span data-stu-id="cff5c-319">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cff5c-320">パラメーター</span><span class="sxs-lookup"><span data-stu-id="cff5c-320">Parameters</span></span>

|<span data-ttu-id="cff5c-321">名前</span><span class="sxs-lookup"><span data-stu-id="cff5c-321">Name</span></span>| <span data-ttu-id="cff5c-322">型</span><span class="sxs-lookup"><span data-stu-id="cff5c-322">Type</span></span>| <span data-ttu-id="cff5c-323">説明</span><span class="sxs-lookup"><span data-stu-id="cff5c-323">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="cff5c-324">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="cff5c-324">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime)|<span data-ttu-id="cff5c-325">変換するローカル時刻の値。</span><span class="sxs-lookup"><span data-stu-id="cff5c-325">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cff5c-326">要件</span><span class="sxs-lookup"><span data-stu-id="cff5c-326">Requirements</span></span>

|<span data-ttu-id="cff5c-327">要件</span><span class="sxs-lookup"><span data-stu-id="cff5c-327">Requirement</span></span>| <span data-ttu-id="cff5c-328">値</span><span class="sxs-lookup"><span data-stu-id="cff5c-328">Value</span></span>|
|---|---|
|[<span data-ttu-id="cff5c-329">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cff5c-329">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cff5c-330">1.0</span><span class="sxs-lookup"><span data-stu-id="cff5c-330">1.0</span></span>|
|[<span data-ttu-id="cff5c-331">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cff5c-331">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cff5c-332">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cff5c-332">ReadItem</span></span>|
|[<span data-ttu-id="cff5c-333">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cff5c-333">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cff5c-334">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="cff5c-334">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="cff5c-335">戻り値:</span><span class="sxs-lookup"><span data-stu-id="cff5c-335">Returns:</span></span>

<span data-ttu-id="cff5c-336">時間が UTC で表現された日付オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="cff5c-336">A Date object with the time expressed in UTC.</span></span>

<span data-ttu-id="cff5c-337">型: Date</span><span class="sxs-lookup"><span data-stu-id="cff5c-337">Type: Date</span></span>

##### <a name="example"></a><span data-ttu-id="cff5c-338">例</span><span class="sxs-lookup"><span data-stu-id="cff5c-338">Example</span></span>

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

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="cff5c-339">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="cff5c-339">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="cff5c-340">既存の予定を表示します。</span><span class="sxs-lookup"><span data-stu-id="cff5c-340">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="cff5c-341">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="cff5c-341">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="cff5c-342">`displayAppointmentForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存の予定を開きます。</span><span class="sxs-lookup"><span data-stu-id="cff5c-342">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="cff5c-p110">Outlook on Mac では、このメソッドを使用して定期的な系列に含まれない単発の予定や定期的な系列のマスター予定を表示できます。ただし、系列のインスタンスは表示できません。これは、Outlook on Mac では定期的な系列のインスタンスのプロパティ (アイテム ID を含む) にアクセスできないためです。</span><span class="sxs-lookup"><span data-stu-id="cff5c-p110">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="cff5c-345">Outlook on the web では、このメソッドはフォームの本文が 32KB 以下の文字数の場合にのみ指定のフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="cff5c-345">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="cff5c-346">指定のアイテム識別子が既存の予定を表していない場合は、クライアント コンピューターまたはデバイスで空のウィンドウが開き、エラー メッセージは返されません。</span><span class="sxs-lookup"><span data-stu-id="cff5c-346">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cff5c-347">パラメーター</span><span class="sxs-lookup"><span data-stu-id="cff5c-347">Parameters</span></span>

|<span data-ttu-id="cff5c-348">名前</span><span class="sxs-lookup"><span data-stu-id="cff5c-348">Name</span></span>| <span data-ttu-id="cff5c-349">型</span><span class="sxs-lookup"><span data-stu-id="cff5c-349">Type</span></span>| <span data-ttu-id="cff5c-350">説明</span><span class="sxs-lookup"><span data-stu-id="cff5c-350">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="cff5c-351">String</span><span class="sxs-lookup"><span data-stu-id="cff5c-351">String</span></span>|<span data-ttu-id="cff5c-352">既存の予定の Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="cff5c-352">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cff5c-353">要件</span><span class="sxs-lookup"><span data-stu-id="cff5c-353">Requirements</span></span>

|<span data-ttu-id="cff5c-354">要件</span><span class="sxs-lookup"><span data-stu-id="cff5c-354">Requirement</span></span>| <span data-ttu-id="cff5c-355">値</span><span class="sxs-lookup"><span data-stu-id="cff5c-355">Value</span></span>|
|---|---|
|[<span data-ttu-id="cff5c-356">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cff5c-356">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cff5c-357">1.0</span><span class="sxs-lookup"><span data-stu-id="cff5c-357">1.0</span></span>|
|[<span data-ttu-id="cff5c-358">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cff5c-358">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cff5c-359">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cff5c-359">ReadItem</span></span>|
|[<span data-ttu-id="cff5c-360">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cff5c-360">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cff5c-361">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="cff5c-361">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cff5c-362">例</span><span class="sxs-lookup"><span data-stu-id="cff5c-362">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

<br>

---
---

#### <a name="displaymessageformitemid"></a><span data-ttu-id="cff5c-363">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="cff5c-363">displayMessageForm(itemId)</span></span>

<span data-ttu-id="cff5c-364">既存のメッセージを表示します。</span><span class="sxs-lookup"><span data-stu-id="cff5c-364">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="cff5c-365">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="cff5c-365">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="cff5c-366">`displayMessageForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存のメッセージを開きます。</span><span class="sxs-lookup"><span data-stu-id="cff5c-366">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="cff5c-367">Outlook on the web では、このメソッドはフォームの本文が 32 KB 以下の文字数の場合にのみ指定のフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="cff5c-367">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="cff5c-368">指定のアイテム識別子が既存のメッセージを表していない場合は、クライアント コンピューターにはメッセージは表示されず、エラー メッセージも返されません。</span><span class="sxs-lookup"><span data-stu-id="cff5c-368">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="cff5c-p111">予定を表す `itemId` を含む `displayMessageForm` を使用しないでください。既存の予定を表示するには、`displayAppointmentForm` メソッドを使用します。新しい予定を作成するフォームを表示するには、`displayNewAppointmentForm` メソッドを使用します。</span><span class="sxs-lookup"><span data-stu-id="cff5c-p111">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cff5c-371">パラメーター</span><span class="sxs-lookup"><span data-stu-id="cff5c-371">Parameters</span></span>

|<span data-ttu-id="cff5c-372">名前</span><span class="sxs-lookup"><span data-stu-id="cff5c-372">Name</span></span>| <span data-ttu-id="cff5c-373">型</span><span class="sxs-lookup"><span data-stu-id="cff5c-373">Type</span></span>| <span data-ttu-id="cff5c-374">説明</span><span class="sxs-lookup"><span data-stu-id="cff5c-374">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="cff5c-375">String</span><span class="sxs-lookup"><span data-stu-id="cff5c-375">String</span></span>|<span data-ttu-id="cff5c-376">既存のメッセージの Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="cff5c-376">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cff5c-377">要件</span><span class="sxs-lookup"><span data-stu-id="cff5c-377">Requirements</span></span>

|<span data-ttu-id="cff5c-378">要件</span><span class="sxs-lookup"><span data-stu-id="cff5c-378">Requirement</span></span>| <span data-ttu-id="cff5c-379">値</span><span class="sxs-lookup"><span data-stu-id="cff5c-379">Value</span></span>|
|---|---|
|[<span data-ttu-id="cff5c-380">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cff5c-380">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cff5c-381">1.0</span><span class="sxs-lookup"><span data-stu-id="cff5c-381">1.0</span></span>|
|[<span data-ttu-id="cff5c-382">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cff5c-382">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cff5c-383">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cff5c-383">ReadItem</span></span>|
|[<span data-ttu-id="cff5c-384">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cff5c-384">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cff5c-385">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="cff5c-385">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cff5c-386">例</span><span class="sxs-lookup"><span data-stu-id="cff5c-386">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

<br>

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="cff5c-387">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="cff5c-387">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="cff5c-388">新しい予定を作成するためのフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="cff5c-388">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="cff5c-389">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="cff5c-389">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="cff5c-p112">`displayNewAppointmentForm` メソッドを使用すると、ユーザーが新しい予定または会議を作成できるフォームが開きます。パラメーターを指定すると、予定のフォーム フィールドにパラメーターの内容が自動的に設定されます。</span><span class="sxs-lookup"><span data-stu-id="cff5c-p112">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="cff5c-p113">Outlook on the web およびモバイル デバイスでは、このメソッドは常に出席者フィールドが含まれるフォームを表示します。入力引数として出席者を指定しないと、このメソッドは **[保存]** ボタンのあるフォームを表示します。出席者を指定した場合には、フォームにその出席者と **[送信]** ボタンが表示されます。</span><span class="sxs-lookup"><span data-stu-id="cff5c-p113">In Outlook on the web and mobile devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="cff5c-p114">Outlook リッチ クライアントと Outlook RT で、`requiredAttendees`、`optionalAttendees`、または `resources` パラメーターに出席者またはリソースを指定し、このメソッドを実行すると、**[送信]** ボタンがある会議フォームが表示されます。受信者を指定せずにこのメソッドを実行すると、**[保存して閉じる]** ボタンがある予定フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="cff5c-p114">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="cff5c-397">パラメーターのいずれかが指定のサイズ制限を超える場合、または不明なパラメーター名が指定されている場合は、例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="cff5c-397">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cff5c-398">パラメーター</span><span class="sxs-lookup"><span data-stu-id="cff5c-398">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="cff5c-399">すべてのパラメーターは省略可能です。</span><span class="sxs-lookup"><span data-stu-id="cff5c-399">All parameters are optional.</span></span>

|<span data-ttu-id="cff5c-400">名前</span><span class="sxs-lookup"><span data-stu-id="cff5c-400">Name</span></span>| <span data-ttu-id="cff5c-401">型</span><span class="sxs-lookup"><span data-stu-id="cff5c-401">Type</span></span>| <span data-ttu-id="cff5c-402">説明</span><span class="sxs-lookup"><span data-stu-id="cff5c-402">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="cff5c-403">Object</span><span class="sxs-lookup"><span data-stu-id="cff5c-403">Object</span></span> | <span data-ttu-id="cff5c-404">新しい予定を記述するパラメーターのディクショナリ。</span><span class="sxs-lookup"><span data-stu-id="cff5c-404">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="cff5c-405">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="cff5c-405">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="cff5c-p115">予定に必要な各出席者について、メール アドレスを含む文字列の配列、または `EmailAddressDetails` オブジェクトを含む配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="cff5c-p115">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="cff5c-408">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="cff5c-408">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="cff5c-p116">予定の各任意出席者について、メール アドレスを含む文字列の配列、または `EmailAddressDetails` オブジェクトを含む配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="cff5c-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="cff5c-411">日付</span><span class="sxs-lookup"><span data-stu-id="cff5c-411">Date</span></span> | <span data-ttu-id="cff5c-412">予定の開始日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="cff5c-412">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="cff5c-413">日付</span><span class="sxs-lookup"><span data-stu-id="cff5c-413">Date</span></span> | <span data-ttu-id="cff5c-414">予定の終了日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="cff5c-414">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="cff5c-415">String</span><span class="sxs-lookup"><span data-stu-id="cff5c-415">String</span></span> | <span data-ttu-id="cff5c-p117">予定の場所を含む文字列。文字列は最大 255 文字に制限されます。</span><span class="sxs-lookup"><span data-stu-id="cff5c-p117">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="cff5c-418">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="cff5c-418">Array.&lt;String&gt;</span></span> | <span data-ttu-id="cff5c-p118">予定に必要なリソースを含む文字列の配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="cff5c-p118">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="cff5c-421">String</span><span class="sxs-lookup"><span data-stu-id="cff5c-421">String</span></span> | <span data-ttu-id="cff5c-p119">予定の件名を含む文字列です。文字列は最大 255 文字に制限されます。</span><span class="sxs-lookup"><span data-stu-id="cff5c-p119">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="cff5c-424">String</span><span class="sxs-lookup"><span data-stu-id="cff5c-424">String</span></span> | <span data-ttu-id="cff5c-p120">予定の本文。本文の内容は、最大サイズが 32 KB に制限されます。</span><span class="sxs-lookup"><span data-stu-id="cff5c-p120">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="cff5c-427">要件</span><span class="sxs-lookup"><span data-stu-id="cff5c-427">Requirements</span></span>

|<span data-ttu-id="cff5c-428">要件</span><span class="sxs-lookup"><span data-stu-id="cff5c-428">Requirement</span></span>| <span data-ttu-id="cff5c-429">値</span><span class="sxs-lookup"><span data-stu-id="cff5c-429">Value</span></span>|
|---|---|
|[<span data-ttu-id="cff5c-430">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cff5c-430">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cff5c-431">1.0</span><span class="sxs-lookup"><span data-stu-id="cff5c-431">1.0</span></span>|
|[<span data-ttu-id="cff5c-432">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cff5c-432">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cff5c-433">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cff5c-433">ReadItem</span></span>|
|[<span data-ttu-id="cff5c-434">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cff5c-434">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cff5c-435">読み取り</span><span class="sxs-lookup"><span data-stu-id="cff5c-435">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cff5c-436">例</span><span class="sxs-lookup"><span data-stu-id="cff5c-436">Example</span></span>

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

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="cff5c-437">displayNewMessageForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="cff5c-437">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="cff5c-438">新しいメッセージを作成するためのフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="cff5c-438">Displays a form for creating a new message.</span></span>

<span data-ttu-id="cff5c-p121">The `displayNewMessageForm` method opens a form that enables the user to create a new message. If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span><span class="sxs-lookup"><span data-stu-id="cff5c-p121">The `displayNewMessageForm` method opens a form that enables the user to create a new message. If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="cff5c-441">パラメーターのいずれかが指定のサイズ制限を超える場合、または不明なパラメーター名が指定されている場合は、例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="cff5c-441">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cff5c-442">パラメーター</span><span class="sxs-lookup"><span data-stu-id="cff5c-442">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="cff5c-443">すべてのパラメーターは省略可能です。</span><span class="sxs-lookup"><span data-stu-id="cff5c-443">All parameters are optional.</span></span>

|<span data-ttu-id="cff5c-444">名前</span><span class="sxs-lookup"><span data-stu-id="cff5c-444">Name</span></span>| <span data-ttu-id="cff5c-445">型</span><span class="sxs-lookup"><span data-stu-id="cff5c-445">Type</span></span>| <span data-ttu-id="cff5c-446">説明</span><span class="sxs-lookup"><span data-stu-id="cff5c-446">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="cff5c-447">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="cff5c-447">Object</span></span> | <span data-ttu-id="cff5c-448">新しいメッセージを記述するパラメータの辞書。</span><span class="sxs-lookup"><span data-stu-id="cff5c-448">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="cff5c-449">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="cff5c-449">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="cff5c-p122">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line. The array is limited to a maximum of 100 entries.</span><span class="sxs-lookup"><span data-stu-id="cff5c-p122">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="cff5c-452">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="cff5c-452">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="cff5c-p123">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line. The array is limited to a maximum of 100 entries.</span><span class="sxs-lookup"><span data-stu-id="cff5c-p123">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="cff5c-455">配列。&lt;文字列&gt; | 配列。&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="cff5c-455">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="cff5c-p124">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line. The array is limited to a maximum of 100 entries.</span><span class="sxs-lookup"><span data-stu-id="cff5c-p124">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="cff5c-458">String</span><span class="sxs-lookup"><span data-stu-id="cff5c-458">String</span></span> | <span data-ttu-id="cff5c-p125">A string containing the subject of the message. The string is limited to a maximum of 255 characters.</span><span class="sxs-lookup"><span data-stu-id="cff5c-p125">A string containing the subject of the message. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="cff5c-461">String</span><span class="sxs-lookup"><span data-stu-id="cff5c-461">String</span></span> | <span data-ttu-id="cff5c-p126">The HTML body of the message. The body content is limited to a maximum size of 32 KB.</span><span class="sxs-lookup"><span data-stu-id="cff5c-p126">The HTML body of the message. The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="cff5c-464">配列。&lt;オブジェクト&gt;</span><span class="sxs-lookup"><span data-stu-id="cff5c-464">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="cff5c-465">添付ファイルまたは添付アイテムである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="cff5c-465">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="cff5c-466">String</span><span class="sxs-lookup"><span data-stu-id="cff5c-466">String</span></span> | <span data-ttu-id="cff5c-p127">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="cff5c-p127">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="cff5c-469">String</span><span class="sxs-lookup"><span data-stu-id="cff5c-469">String</span></span> | <span data-ttu-id="cff5c-470">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="cff5c-470">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="cff5c-471">文字列</span><span class="sxs-lookup"><span data-stu-id="cff5c-471">String</span></span> | <span data-ttu-id="cff5c-p128">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="cff5c-p128">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="cff5c-474">ブール値</span><span class="sxs-lookup"><span data-stu-id="cff5c-474">Boolean</span></span> | <span data-ttu-id="cff5c-p129">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="cff5c-p129">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="cff5c-477">String</span><span class="sxs-lookup"><span data-stu-id="cff5c-477">String</span></span> | <span data-ttu-id="cff5c-p130">Only used if `type` is set to `item`. The EWS item id of the existing e-mail you want to attach to the new message. This is a string up to 100 characters.</span><span class="sxs-lookup"><span data-stu-id="cff5c-p130">Only used if `type` is set to `item`. The EWS item id of the existing e-mail you want to attach to the new message. This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="cff5c-481">要件</span><span class="sxs-lookup"><span data-stu-id="cff5c-481">Requirements</span></span>

|<span data-ttu-id="cff5c-482">要件</span><span class="sxs-lookup"><span data-stu-id="cff5c-482">Requirement</span></span>| <span data-ttu-id="cff5c-483">値</span><span class="sxs-lookup"><span data-stu-id="cff5c-483">Value</span></span>|
|---|---|
|[<span data-ttu-id="cff5c-484">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cff5c-484">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cff5c-485">1.6</span><span class="sxs-lookup"><span data-stu-id="cff5c-485">1.6</span></span> |
|[<span data-ttu-id="cff5c-486">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cff5c-486">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cff5c-487">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cff5c-487">ReadItem</span></span>|
|[<span data-ttu-id="cff5c-488">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cff5c-488">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cff5c-489">読み取り</span><span class="sxs-lookup"><span data-stu-id="cff5c-489">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cff5c-490">例</span><span class="sxs-lookup"><span data-stu-id="cff5c-490">Example</span></span>

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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="cff5c-491">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="cff5c-491">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="cff5c-492">REST API または Exchange Web サービスを呼び出すために使用するトークンを含む文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="cff5c-492">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="cff5c-p131">`getCallbackTokenAsync` メソッドは、ユーザーのメールボックスをホストする Exchange Server から不透明なトークンを取得する非同期の呼び出しを行います。コールバック トークンの有効期間は 5 分です。</span><span class="sxs-lookup"><span data-stu-id="cff5c-p131">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="cff5c-495">可能であれば、アドインでは Exchange Web サービスの代わりに REST API を使用することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="cff5c-495">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span>

<span data-ttu-id="cff5c-496">閲覧モードで `getCallbackTokenAsync` メソッドを呼び出すには、**ReadItem** の最小限のアクセス許可レベルが必要です。</span><span class="sxs-lookup"><span data-stu-id="cff5c-496">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="cff5c-497">作成モードで `getCallbackTokenAsync` を呼び出すには、アイテムを保存しておく必要があります。</span><span class="sxs-lookup"><span data-stu-id="cff5c-497">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="cff5c-498">[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドは、**ReadWriteItem** の最小限のアクセス許可レベルが必要です。</span><span class="sxs-lookup"><span data-stu-id="cff5c-498">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

<span data-ttu-id="cff5c-499">**REST トークン**</span><span class="sxs-lookup"><span data-stu-id="cff5c-499">**REST Tokens**</span></span>

<span data-ttu-id="cff5c-p133">REST トークンが要求された場合 (`options.isRest = true`)、結果トークンは Exchange Web サービスの呼び出しを認証するためには機能しません。アドインがマニフェストで [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) アクセス許可を指定していない限り、トークンの範囲は現在のアイテムとその添付ファイルへの読み取り専用アクセスに制限されます。`ReadWriteMailbox` アクセス許可が指定されている場合は、結果トークンは、メールを送信する機能など、メール、カレンダー、連絡先への読み取り/書き込みアクセスを付与します。</span><span class="sxs-lookup"><span data-stu-id="cff5c-p133">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="cff5c-503">アドインでは、`restUrl` プロパティを使用して、REST API 呼び出しを行うときに使用する正しい URL を決定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="cff5c-503">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="cff5c-504">**EWS トークン**</span><span class="sxs-lookup"><span data-stu-id="cff5c-504">**EWS Tokens**</span></span>

<span data-ttu-id="cff5c-p134">EWS トークンが要求された場合 (`options.isRest = false`)、結果トークンは REST API 呼び出しを認証するためには機能しません。トークンの範囲は、現在のアイテムへのアクセスに制限されます。</span><span class="sxs-lookup"><span data-stu-id="cff5c-p134">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="cff5c-507">アドインでは、`ewsUrl` プロパティを使用して、EWS 呼び出しを行うときに使用する正しい URL を決定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="cff5c-507">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

<span data-ttu-id="cff5c-508">トークンと、添付ファイル識別子またはアイテム識別子の両方をサードパーティ システムに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="cff5c-508">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="cff5c-509">サードパーティシステムは、トークンをベアラー認証トークンとして使用して、Exchange Web サービス (EWS) の[Getattachment](/exchange/client-developer/web-service-reference/getattachment-operation)操作または[GetItem](/exchange/client-developer/web-service-reference/getitem-operation)操作を呼び出して、添付ファイルまたはアイテムを取得します。</span><span class="sxs-lookup"><span data-stu-id="cff5c-509">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to retrieve an attachment or item.</span></span> <span data-ttu-id="cff5c-510">たとえば、[選択したアイテムから添付ファイルを取得する](/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="cff5c-510">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

##### <a name="parameters"></a><span data-ttu-id="cff5c-511">パラメーター</span><span class="sxs-lookup"><span data-stu-id="cff5c-511">Parameters</span></span>

|<span data-ttu-id="cff5c-512">名前</span><span class="sxs-lookup"><span data-stu-id="cff5c-512">Name</span></span>| <span data-ttu-id="cff5c-513">型</span><span class="sxs-lookup"><span data-stu-id="cff5c-513">Type</span></span>| <span data-ttu-id="cff5c-514">属性</span><span class="sxs-lookup"><span data-stu-id="cff5c-514">Attributes</span></span>| <span data-ttu-id="cff5c-515">説明</span><span class="sxs-lookup"><span data-stu-id="cff5c-515">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="cff5c-516">Object</span><span class="sxs-lookup"><span data-stu-id="cff5c-516">Object</span></span> | <span data-ttu-id="cff5c-517">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="cff5c-517">&lt;optional&gt;</span></span> | <span data-ttu-id="cff5c-518">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="cff5c-518">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="cff5c-519">ブール値</span><span class="sxs-lookup"><span data-stu-id="cff5c-519">Boolean</span></span> |  <span data-ttu-id="cff5c-520">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="cff5c-520">&lt;optional&gt;</span></span> | <span data-ttu-id="cff5c-p136">提供されたトークンを Outlook REST API または Exchange Web サービスに使用するかどうかを決定します。既定値は、`false` です。</span><span class="sxs-lookup"><span data-stu-id="cff5c-p136">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="cff5c-523">Object</span><span class="sxs-lookup"><span data-stu-id="cff5c-523">Object</span></span> |  <span data-ttu-id="cff5c-524">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="cff5c-524">&lt;optional&gt;</span></span> | <span data-ttu-id="cff5c-525">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="cff5c-525">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="cff5c-526">function</span><span class="sxs-lookup"><span data-stu-id="cff5c-526">function</span></span>||<span data-ttu-id="cff5c-527">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="cff5c-527">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="cff5c-528">トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="cff5c-528">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="cff5c-529">エラーが発生した場合、 `asyncResult.error` および `asyncResult.diagnostics` のプロパティで追加情報が提供される場合があります。</span><span class="sxs-lookup"><span data-stu-id="cff5c-529">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|

##### <a name="errors"></a><span data-ttu-id="cff5c-530">エラー</span><span class="sxs-lookup"><span data-stu-id="cff5c-530">Errors</span></span>

|<span data-ttu-id="cff5c-531">エラー コード</span><span class="sxs-lookup"><span data-stu-id="cff5c-531">Error code</span></span>|<span data-ttu-id="cff5c-532">説明</span><span class="sxs-lookup"><span data-stu-id="cff5c-532">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="cff5c-533">要求が失敗しました。</span><span class="sxs-lookup"><span data-stu-id="cff5c-533">The request has failed.</span></span> <span data-ttu-id="cff5c-534">HTTP エラーコードの diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="cff5c-534">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="cff5c-535">Exchange サーバーがエラーを返しました。</span><span class="sxs-lookup"><span data-stu-id="cff5c-535">The Exchange server returned an error.</span></span> <span data-ttu-id="cff5c-536">詳細については、diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="cff5c-536">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="cff5c-537">ユーザーはネットワークに接続されていません。</span><span class="sxs-lookup"><span data-stu-id="cff5c-537">The user is no longer connected to the network.</span></span> <span data-ttu-id="cff5c-538">ネットワーク接続を確認し、やり直してください。</span><span class="sxs-lookup"><span data-stu-id="cff5c-538">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cff5c-539">要件</span><span class="sxs-lookup"><span data-stu-id="cff5c-539">Requirements</span></span>

|<span data-ttu-id="cff5c-540">必要条件</span><span class="sxs-lookup"><span data-stu-id="cff5c-540">Requirement</span></span>| <span data-ttu-id="cff5c-541">値</span><span class="sxs-lookup"><span data-stu-id="cff5c-541">Value</span></span>|
|---|---|
|[<span data-ttu-id="cff5c-542">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cff5c-542">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cff5c-543">1.5</span><span class="sxs-lookup"><span data-stu-id="cff5c-543">1.5</span></span> |
|[<span data-ttu-id="cff5c-544">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cff5c-544">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cff5c-545">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cff5c-545">ReadItem</span></span>|
|[<span data-ttu-id="cff5c-546">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cff5c-546">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cff5c-547">新規作成と閲覧</span><span class="sxs-lookup"><span data-stu-id="cff5c-547">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="cff5c-548">例</span><span class="sxs-lookup"><span data-stu-id="cff5c-548">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="cff5c-549">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="cff5c-549">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="cff5c-550">Exchange Server から添付ファイルやアイテムを取得するために使うトークンを含む文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="cff5c-550">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="cff5c-p140">`getCallbackTokenAsync` メソッドは、ユーザーのメールボックスをホストする Exchange Server から不透明なトークンを取得する非同期の呼び出しを行います。コールバック トークンの有効期間は 5 分です。</span><span class="sxs-lookup"><span data-stu-id="cff5c-p140">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="cff5c-553">トークンと、添付ファイル識別子またはアイテム識別子の両方をサードパーティ システムに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="cff5c-553">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="cff5c-554">サードパーティ システムは、トークンをベアラー承認トークンとして使用し、Exchange Web サービス (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) 操作または [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) 操作を呼び出して、添付ファイルまたはアイテムを返します。</span><span class="sxs-lookup"><span data-stu-id="cff5c-554">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item.</span></span> <span data-ttu-id="cff5c-555">たとえば、[選択したアイテムから添付ファイルを取得する](/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="cff5c-555">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="cff5c-556">閲覧モードで `getCallbackTokenAsync` メソッドを呼び出すには、**ReadItem** の最小限のアクセス許可レベルが必要です。</span><span class="sxs-lookup"><span data-stu-id="cff5c-556">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="cff5c-557">作成モードで `getCallbackTokenAsync` を呼び出すには、アイテムを保存しておく必要があります。</span><span class="sxs-lookup"><span data-stu-id="cff5c-557">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="cff5c-558">[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドは、**ReadWriteItem** の最小限のアクセス許可レベルが必要です。</span><span class="sxs-lookup"><span data-stu-id="cff5c-558">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cff5c-559">パラメーター</span><span class="sxs-lookup"><span data-stu-id="cff5c-559">Parameters</span></span>

|<span data-ttu-id="cff5c-560">名前</span><span class="sxs-lookup"><span data-stu-id="cff5c-560">Name</span></span>| <span data-ttu-id="cff5c-561">型</span><span class="sxs-lookup"><span data-stu-id="cff5c-561">Type</span></span>| <span data-ttu-id="cff5c-562">属性</span><span class="sxs-lookup"><span data-stu-id="cff5c-562">Attributes</span></span>| <span data-ttu-id="cff5c-563">説明</span><span class="sxs-lookup"><span data-stu-id="cff5c-563">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="cff5c-564">関数</span><span class="sxs-lookup"><span data-stu-id="cff5c-564">function</span></span>||<span data-ttu-id="cff5c-565">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="cff5c-565">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="cff5c-566">トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="cff5c-566">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="cff5c-567">エラーが発生した場合、 `asyncResult.error` および `asyncResult.diagnostics` のプロパティで追加情報が提供される場合があります。</span><span class="sxs-lookup"><span data-stu-id="cff5c-567">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="cff5c-568">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="cff5c-568">Object</span></span>| <span data-ttu-id="cff5c-569">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="cff5c-569">&lt;optional&gt;</span></span>|<span data-ttu-id="cff5c-570">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="cff5c-570">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="cff5c-571">エラー</span><span class="sxs-lookup"><span data-stu-id="cff5c-571">Errors</span></span>

|<span data-ttu-id="cff5c-572">エラー コード</span><span class="sxs-lookup"><span data-stu-id="cff5c-572">Error code</span></span>|<span data-ttu-id="cff5c-573">説明</span><span class="sxs-lookup"><span data-stu-id="cff5c-573">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="cff5c-574">要求が失敗しました。</span><span class="sxs-lookup"><span data-stu-id="cff5c-574">The request has failed.</span></span> <span data-ttu-id="cff5c-575">HTTP エラーコードの diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="cff5c-575">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="cff5c-576">Exchange サーバーがエラーを返しました。</span><span class="sxs-lookup"><span data-stu-id="cff5c-576">The Exchange server returned an error.</span></span> <span data-ttu-id="cff5c-577">詳細については、diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="cff5c-577">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="cff5c-578">ユーザーはネットワークに接続されていません。</span><span class="sxs-lookup"><span data-stu-id="cff5c-578">The user is no longer connected to the network.</span></span> <span data-ttu-id="cff5c-579">ネットワーク接続を確認し、やり直してください。</span><span class="sxs-lookup"><span data-stu-id="cff5c-579">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cff5c-580">要件</span><span class="sxs-lookup"><span data-stu-id="cff5c-580">Requirements</span></span>

|<span data-ttu-id="cff5c-581">要件</span><span class="sxs-lookup"><span data-stu-id="cff5c-581">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="cff5c-582">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cff5c-582">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cff5c-583">1.0</span><span class="sxs-lookup"><span data-stu-id="cff5c-583">1.0</span></span> | <span data-ttu-id="cff5c-584">1.3</span><span class="sxs-lookup"><span data-stu-id="cff5c-584">1.3</span></span> |
|[<span data-ttu-id="cff5c-585">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cff5c-585">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cff5c-586">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cff5c-586">ReadItem</span></span> | <span data-ttu-id="cff5c-587">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cff5c-587">ReadItem</span></span> |
|[<span data-ttu-id="cff5c-588">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cff5c-588">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cff5c-589">Read</span><span class="sxs-lookup"><span data-stu-id="cff5c-589">Read</span></span> | <span data-ttu-id="cff5c-590">Compose</span><span class="sxs-lookup"><span data-stu-id="cff5c-590">Compose</span></span> |

##### <a name="example"></a><span data-ttu-id="cff5c-591">例</span><span class="sxs-lookup"><span data-stu-id="cff5c-591">Example</span></span>

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

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="cff5c-592">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="cff5c-592">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="cff5c-593">ユーザーと Office アドインを識別するトークンを取得します。</span><span class="sxs-lookup"><span data-stu-id="cff5c-593">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="cff5c-594">`getUserIdentityTokenAsync` メソッドは、[アドインとユーザーをサード パーティのシステムで識別して認証](/outlook/add-ins/authentication)することのできるトークンを返します。</span><span class="sxs-lookup"><span data-stu-id="cff5c-594">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="cff5c-595">パラメーター</span><span class="sxs-lookup"><span data-stu-id="cff5c-595">Parameters</span></span>

|<span data-ttu-id="cff5c-596">名前</span><span class="sxs-lookup"><span data-stu-id="cff5c-596">Name</span></span>| <span data-ttu-id="cff5c-597">型</span><span class="sxs-lookup"><span data-stu-id="cff5c-597">Type</span></span>| <span data-ttu-id="cff5c-598">属性</span><span class="sxs-lookup"><span data-stu-id="cff5c-598">Attributes</span></span>| <span data-ttu-id="cff5c-599">説明</span><span class="sxs-lookup"><span data-stu-id="cff5c-599">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="cff5c-600">関数</span><span class="sxs-lookup"><span data-stu-id="cff5c-600">function</span></span>||<span data-ttu-id="cff5c-601">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="cff5c-601">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="cff5c-602">トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="cff5c-602">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="cff5c-603">エラーが発生した場合、 `asyncResult.error` および `asyncResult.diagnostics` のプロパティで追加情報が提供される場合があります。</span><span class="sxs-lookup"><span data-stu-id="cff5c-603">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="cff5c-604">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="cff5c-604">Object</span></span>| <span data-ttu-id="cff5c-605">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="cff5c-605">&lt;optional&gt;</span></span>|<span data-ttu-id="cff5c-606">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="cff5c-606">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="cff5c-607">エラー</span><span class="sxs-lookup"><span data-stu-id="cff5c-607">Errors</span></span>

|<span data-ttu-id="cff5c-608">エラー コード</span><span class="sxs-lookup"><span data-stu-id="cff5c-608">Error code</span></span>|<span data-ttu-id="cff5c-609">説明</span><span class="sxs-lookup"><span data-stu-id="cff5c-609">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="cff5c-610">要求が失敗しました。</span><span class="sxs-lookup"><span data-stu-id="cff5c-610">The request has failed.</span></span> <span data-ttu-id="cff5c-611">HTTP エラーコードの diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="cff5c-611">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="cff5c-612">Exchange サーバーがエラーを返しました。</span><span class="sxs-lookup"><span data-stu-id="cff5c-612">The Exchange server returned an error.</span></span> <span data-ttu-id="cff5c-613">詳細については、diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="cff5c-613">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="cff5c-614">ユーザーはネットワークに接続されていません。</span><span class="sxs-lookup"><span data-stu-id="cff5c-614">The user is no longer connected to the network.</span></span> <span data-ttu-id="cff5c-615">ネットワーク接続を確認し、やり直してください。</span><span class="sxs-lookup"><span data-stu-id="cff5c-615">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cff5c-616">要件</span><span class="sxs-lookup"><span data-stu-id="cff5c-616">Requirements</span></span>

|<span data-ttu-id="cff5c-617">要件</span><span class="sxs-lookup"><span data-stu-id="cff5c-617">Requirement</span></span>| <span data-ttu-id="cff5c-618">値</span><span class="sxs-lookup"><span data-stu-id="cff5c-618">Value</span></span>|
|---|---|
|[<span data-ttu-id="cff5c-619">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cff5c-619">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cff5c-620">1.0</span><span class="sxs-lookup"><span data-stu-id="cff5c-620">1.0</span></span>|
|[<span data-ttu-id="cff5c-621">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cff5c-621">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cff5c-622">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cff5c-622">ReadItem</span></span>|
|[<span data-ttu-id="cff5c-623">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cff5c-623">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cff5c-624">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="cff5c-624">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cff5c-625">例</span><span class="sxs-lookup"><span data-stu-id="cff5c-625">Example</span></span>

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

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="cff5c-626">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="cff5c-626">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="cff5c-627">ユーザーのメールボックスをホストしている Exchange サーバー上の Exchange Web サービス (EWS) のサービスに対して非同期の要求を行います。</span><span class="sxs-lookup"><span data-stu-id="cff5c-627">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="cff5c-628">このメソッドは、次のシナリオではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="cff5c-628">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="cff5c-629">Outlook on iOS または Android の場合</span><span class="sxs-lookup"><span data-stu-id="cff5c-629">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="cff5c-630">アドインが Gmail のメールボックスに読み込まれる場合</span><span class="sxs-lookup"><span data-stu-id="cff5c-630">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="cff5c-631">このような場合は、アドインでは [REST API を使用](/outlook/add-ins/use-rest-api)して、代わりにユーザーのメールボックスにアクセスする必要があります。</span><span class="sxs-lookup"><span data-stu-id="cff5c-631">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="cff5c-p149">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange. See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span><span class="sxs-lookup"><span data-stu-id="cff5c-p149">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange. See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="cff5c-634">`makeEwsRequestAsync` メソッドでは、フォルダー関連アイテムを要求できません。</span><span class="sxs-lookup"><span data-stu-id="cff5c-634">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="cff5c-635">XML 要求では UTF-8 エンコードを指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="cff5c-635">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="cff5c-p150">`makeEwsRequestAsync` メソッドを使用するには、アドインに **ReadWriteMailbox** アクセス許可が必要です。**ReadWriteMailbox** アクセス許可と、`makeEwsRequestAsync` メソッドで呼び出せる EWS 操作の使い方については、「[ユーザーのメールボックスへのメール アドイン アクセスのアクセス許可を指定する](/outlook/add-ins/understanding-outlook-add-in-permissions)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="cff5c-p150">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="cff5c-638">サーバー管理者は、クライアント アクセス サーバーの EWS ディレクトリで `OAuthAuthentication` を true に設定して、`makeEwsRequestAsync` メソッドで EWS 要求を行うことができるようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="cff5c-638">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="cff5c-639">バージョンの相違点</span><span class="sxs-lookup"><span data-stu-id="cff5c-639">Version differences</span></span>

<span data-ttu-id="cff5c-640">バージョン 15.0.4535.1004 より前のバージョンの Outlook で実行しているメール アプリで `makeEwsRequestAsync` メソッドを使用する場合には、エンコード値を `ISO-8859-1` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="cff5c-640">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="cff5c-641">Outlook on the web でメール アプリを実行している場合は、エンコード値を設定する必要はありません。</span><span class="sxs-lookup"><span data-stu-id="cff5c-641">You do not need to set the encoding value when your mail app is running in Outlook on the web.</span></span> <span data-ttu-id="cff5c-642">メールアプリが web 上の Outlook またはデスクトップクライアントで実行されているかどうかは、mailbox プロパティを使用して判断できます。</span><span class="sxs-lookup"><span data-stu-id="cff5c-642">You can determine whether your mail app is running in Outlook on the web or a desktop client by using the mailbox.diagnostics.hostName property.</span></span> <span data-ttu-id="cff5c-643">mailbox.diagnostics.hostVersion プロパティを使って、どのバージョンの Outlook を使って実行しているかを確認できます。</span><span class="sxs-lookup"><span data-stu-id="cff5c-643">You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cff5c-644">パラメーター</span><span class="sxs-lookup"><span data-stu-id="cff5c-644">Parameters</span></span>

|<span data-ttu-id="cff5c-645">名前</span><span class="sxs-lookup"><span data-stu-id="cff5c-645">Name</span></span>| <span data-ttu-id="cff5c-646">型</span><span class="sxs-lookup"><span data-stu-id="cff5c-646">Type</span></span>| <span data-ttu-id="cff5c-647">属性</span><span class="sxs-lookup"><span data-stu-id="cff5c-647">Attributes</span></span>| <span data-ttu-id="cff5c-648">説明</span><span class="sxs-lookup"><span data-stu-id="cff5c-648">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="cff5c-649">String</span><span class="sxs-lookup"><span data-stu-id="cff5c-649">String</span></span>||<span data-ttu-id="cff5c-650">EWS 要求です。</span><span class="sxs-lookup"><span data-stu-id="cff5c-650">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="cff5c-651">function</span><span class="sxs-lookup"><span data-stu-id="cff5c-651">function</span></span>||<span data-ttu-id="cff5c-652">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="cff5c-652">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="cff5c-p152">The XML result of the EWS call is provided as a string in the `asyncResult.value` property. If the result exceeds 1 MB in size, an error message is returned instead.</span><span class="sxs-lookup"><span data-stu-id="cff5c-p152">The XML result of the EWS call is provided as a string in the `asyncResult.value` property. If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="cff5c-655">Object</span><span class="sxs-lookup"><span data-stu-id="cff5c-655">Object</span></span>| <span data-ttu-id="cff5c-656">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="cff5c-656">&lt;optional&gt;</span></span>|<span data-ttu-id="cff5c-657">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="cff5c-657">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cff5c-658">要件</span><span class="sxs-lookup"><span data-stu-id="cff5c-658">Requirements</span></span>

|<span data-ttu-id="cff5c-659">要件</span><span class="sxs-lookup"><span data-stu-id="cff5c-659">Requirement</span></span>| <span data-ttu-id="cff5c-660">値</span><span class="sxs-lookup"><span data-stu-id="cff5c-660">Value</span></span>|
|---|---|
|[<span data-ttu-id="cff5c-661">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cff5c-661">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cff5c-662">1.0</span><span class="sxs-lookup"><span data-stu-id="cff5c-662">1.0</span></span>|
|[<span data-ttu-id="cff5c-663">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cff5c-663">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cff5c-664">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="cff5c-664">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="cff5c-665">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cff5c-665">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cff5c-666">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="cff5c-666">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cff5c-667">例</span><span class="sxs-lookup"><span data-stu-id="cff5c-667">Example</span></span>

<span data-ttu-id="cff5c-668">次の例は、`GetItem` 操作を使ってアイテムの件名を取得するため、`makeEwsRequestAsync` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="cff5c-668">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="cff5c-669">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="cff5c-669">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="cff5c-670">サポートされているイベントの種類のイベント ハンドラーを削除します。</span><span class="sxs-lookup"><span data-stu-id="cff5c-670">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="cff5c-671">現在、サポートされている`Office.EventType.ItemChanged`イベント`Office.EventType.OfficeThemeChanged`の種類はおよびです。</span><span class="sxs-lookup"><span data-stu-id="cff5c-671">Currently, the supported event types are `Office.EventType.ItemChanged` and `Office.EventType.OfficeThemeChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cff5c-672">パラメーター</span><span class="sxs-lookup"><span data-stu-id="cff5c-672">Parameters</span></span>

| <span data-ttu-id="cff5c-673">名前</span><span class="sxs-lookup"><span data-stu-id="cff5c-673">Name</span></span> | <span data-ttu-id="cff5c-674">型</span><span class="sxs-lookup"><span data-stu-id="cff5c-674">Type</span></span> | <span data-ttu-id="cff5c-675">属性</span><span class="sxs-lookup"><span data-stu-id="cff5c-675">Attributes</span></span> | <span data-ttu-id="cff5c-676">説明</span><span class="sxs-lookup"><span data-stu-id="cff5c-676">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="cff5c-677">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="cff5c-677">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="cff5c-678">ハンドラーを取り消すイベント。</span><span class="sxs-lookup"><span data-stu-id="cff5c-678">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="cff5c-679">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="cff5c-679">Object</span></span> | <span data-ttu-id="cff5c-680">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="cff5c-680">&lt;optional&gt;</span></span> | <span data-ttu-id="cff5c-681">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="cff5c-681">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="cff5c-682">Object</span><span class="sxs-lookup"><span data-stu-id="cff5c-682">Object</span></span> | <span data-ttu-id="cff5c-683">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="cff5c-683">&lt;optional&gt;</span></span> | <span data-ttu-id="cff5c-684">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="cff5c-684">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="cff5c-685">function</span><span class="sxs-lookup"><span data-stu-id="cff5c-685">function</span></span>| <span data-ttu-id="cff5c-686">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="cff5c-686">&lt;optional&gt;</span></span>|<span data-ttu-id="cff5c-687">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="cff5c-687">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cff5c-688">要件</span><span class="sxs-lookup"><span data-stu-id="cff5c-688">Requirements</span></span>

|<span data-ttu-id="cff5c-689">要件</span><span class="sxs-lookup"><span data-stu-id="cff5c-689">Requirement</span></span>| <span data-ttu-id="cff5c-690">値</span><span class="sxs-lookup"><span data-stu-id="cff5c-690">Value</span></span>|
|---|---|
|[<span data-ttu-id="cff5c-691">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cff5c-691">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cff5c-692">1.5</span><span class="sxs-lookup"><span data-stu-id="cff5c-692">1.5</span></span> |
|[<span data-ttu-id="cff5c-693">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cff5c-693">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cff5c-694">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cff5c-694">ReadItem</span></span> |
|[<span data-ttu-id="cff5c-695">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cff5c-695">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cff5c-696">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="cff5c-696">Compose or Read</span></span>|
