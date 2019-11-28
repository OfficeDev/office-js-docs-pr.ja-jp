---
title: Office. メールボックス要件セット1.8
description: ''
ms.date: 11/27/2019
localization_priority: Normal
ms.openlocfilehash: 908eff7b34e63b62fbe250f1a6f810be69b17627
ms.sourcegitcommit: 05a883a7fd89136301ce35aabc57638e9f563288
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/27/2019
ms.locfileid: "39629217"
---
# <a name="mailbox"></a><span data-ttu-id="d1b7e-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="d1b7e-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="d1b7e-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="d1b7e-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="d1b7e-104">Microsoft Outlook の Outlook アドイン オブジェクト モデルへのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d1b7e-105">要件</span><span class="sxs-lookup"><span data-stu-id="d1b7e-105">Requirements</span></span>

|<span data-ttu-id="d1b7e-106">要件</span><span class="sxs-lookup"><span data-stu-id="d1b7e-106">Requirement</span></span>| <span data-ttu-id="d1b7e-107">値</span><span class="sxs-lookup"><span data-stu-id="d1b7e-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="d1b7e-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d1b7e-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d1b7e-109">1.0</span><span class="sxs-lookup"><span data-stu-id="d1b7e-109">1.0</span></span>|
|[<span data-ttu-id="d1b7e-110">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d1b7e-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d1b7e-111">制限あり</span><span class="sxs-lookup"><span data-stu-id="d1b7e-111">Restricted</span></span>|
|[<span data-ttu-id="d1b7e-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d1b7e-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d1b7e-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d1b7e-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="d1b7e-114">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="d1b7e-114">Members and methods</span></span>

| <span data-ttu-id="d1b7e-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="d1b7e-115">Member</span></span> | <span data-ttu-id="d1b7e-116">種類</span><span class="sxs-lookup"><span data-stu-id="d1b7e-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="d1b7e-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="d1b7e-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="d1b7e-118">メンバー</span><span class="sxs-lookup"><span data-stu-id="d1b7e-118">Member</span></span> |
| [<span data-ttu-id="d1b7e-119">masterCategories</span><span class="sxs-lookup"><span data-stu-id="d1b7e-119">masterCategories</span></span>](#mastercategories-mastercategories) | <span data-ttu-id="d1b7e-120">メンバー</span><span class="sxs-lookup"><span data-stu-id="d1b7e-120">Member</span></span> |
| [<span data-ttu-id="d1b7e-121">restUrl</span><span class="sxs-lookup"><span data-stu-id="d1b7e-121">restUrl</span></span>](#resturl-string) | <span data-ttu-id="d1b7e-122">メンバー</span><span class="sxs-lookup"><span data-stu-id="d1b7e-122">Member</span></span> |
| [<span data-ttu-id="d1b7e-123">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="d1b7e-123">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="d1b7e-124">メソッド</span><span class="sxs-lookup"><span data-stu-id="d1b7e-124">Method</span></span> |
| [<span data-ttu-id="d1b7e-125">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="d1b7e-125">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="d1b7e-126">メソッド</span><span class="sxs-lookup"><span data-stu-id="d1b7e-126">Method</span></span> |
| [<span data-ttu-id="d1b7e-127">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="d1b7e-127">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="d1b7e-128">メソッド</span><span class="sxs-lookup"><span data-stu-id="d1b7e-128">Method</span></span> |
| [<span data-ttu-id="d1b7e-129">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="d1b7e-129">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="d1b7e-130">メソッド</span><span class="sxs-lookup"><span data-stu-id="d1b7e-130">Method</span></span> |
| [<span data-ttu-id="d1b7e-131">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="d1b7e-131">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="d1b7e-132">メソッド</span><span class="sxs-lookup"><span data-stu-id="d1b7e-132">Method</span></span> |
| [<span data-ttu-id="d1b7e-133">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="d1b7e-133">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="d1b7e-134">メソッド</span><span class="sxs-lookup"><span data-stu-id="d1b7e-134">Method</span></span> |
| [<span data-ttu-id="d1b7e-135">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="d1b7e-135">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="d1b7e-136">メソッド</span><span class="sxs-lookup"><span data-stu-id="d1b7e-136">Method</span></span> |
| [<span data-ttu-id="d1b7e-137">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="d1b7e-137">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="d1b7e-138">メソッド</span><span class="sxs-lookup"><span data-stu-id="d1b7e-138">Method</span></span> |
| [<span data-ttu-id="d1b7e-139">displayNewMessageForm</span><span class="sxs-lookup"><span data-stu-id="d1b7e-139">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="d1b7e-140">メソッド</span><span class="sxs-lookup"><span data-stu-id="d1b7e-140">Method</span></span> |
| [<span data-ttu-id="d1b7e-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="d1b7e-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="d1b7e-142">メソッド</span><span class="sxs-lookup"><span data-stu-id="d1b7e-142">Method</span></span> |
| [<span data-ttu-id="d1b7e-143">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="d1b7e-143">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="d1b7e-144">メソッド</span><span class="sxs-lookup"><span data-stu-id="d1b7e-144">Method</span></span> |
| [<span data-ttu-id="d1b7e-145">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="d1b7e-145">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="d1b7e-146">メソッド</span><span class="sxs-lookup"><span data-stu-id="d1b7e-146">Method</span></span> |
| [<span data-ttu-id="d1b7e-147">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="d1b7e-147">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="d1b7e-148">メソッド</span><span class="sxs-lookup"><span data-stu-id="d1b7e-148">Method</span></span> |
| [<span data-ttu-id="d1b7e-149">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="d1b7e-149">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="d1b7e-150">メソッド</span><span class="sxs-lookup"><span data-stu-id="d1b7e-150">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="d1b7e-151">名前空間</span><span class="sxs-lookup"><span data-stu-id="d1b7e-151">Namespaces</span></span>

<span data-ttu-id="d1b7e-152">[diagnostics](Office.context.mailbox.diagnostics.md):Outlook アドインに診断情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-152">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="d1b7e-153">[item](Office.context.mailbox.item.md):Outlook アドインのメッセージや予定にアクセスするためのメソッドとプロパティを提供します。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-153">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="d1b7e-154">[userProfile](Office.context.mailbox.userProfile.md):Outlook アドインのユーザーに関する情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-154">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="d1b7e-155">Members</span><span class="sxs-lookup"><span data-stu-id="d1b7e-155">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="d1b7e-156">ewsUrl: String</span><span class="sxs-lookup"><span data-stu-id="d1b7e-156">ewsUrl: String</span></span>

<span data-ttu-id="d1b7e-p101">このメール アカウントの Exchange Web サービス (EWS) エンドポイントの URL を取得します。読み取りモードのみです。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="d1b7e-159">このメンバーは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-159">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="d1b7e-p102">
  `ewsUrl\` 値は、リモート サービスで、ユーザーのメールボックスに EWS 呼び出しを行うために使うことができます。たとえば、[選択したアイテムから添付ファイルを取得する](/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="d1b7e-162">アプリが閲覧モードで `ewsUrl` メンバーを呼び出すには、アプリのマニフェスト内に **ReadItem** アクセス許可が指定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-162">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="d1b7e-p103">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出してから、`ewsUrl` メンバーを使用する必要があります。アプリには、`saveAsync` メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="d1b7e-165">型</span><span class="sxs-lookup"><span data-stu-id="d1b7e-165">Type</span></span>

*   <span data-ttu-id="d1b7e-166">String</span><span class="sxs-lookup"><span data-stu-id="d1b7e-166">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d1b7e-167">Requirements</span><span class="sxs-lookup"><span data-stu-id="d1b7e-167">Requirements</span></span>

|<span data-ttu-id="d1b7e-168">要件</span><span class="sxs-lookup"><span data-stu-id="d1b7e-168">Requirement</span></span>| <span data-ttu-id="d1b7e-169">値</span><span class="sxs-lookup"><span data-stu-id="d1b7e-169">Value</span></span>|
|---|---|
|[<span data-ttu-id="d1b7e-170">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d1b7e-170">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d1b7e-171">1.0</span><span class="sxs-lookup"><span data-stu-id="d1b7e-171">1.0</span></span>|
|[<span data-ttu-id="d1b7e-172">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d1b7e-172">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d1b7e-173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d1b7e-173">ReadItem</span></span>|
|[<span data-ttu-id="d1b7e-174">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d1b7e-174">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d1b7e-175">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d1b7e-175">Compose or Read</span></span>|

<br>

---
---

#### <a name="mastercategories-mastercategoriesjavascriptapioutlookofficemastercategoriesviewoutlook-js-18"></a><span data-ttu-id="d1b7e-176">masterCategories: [Mastercategories](/javascript/api/outlook/office.mastercategories?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="d1b7e-176">masterCategories: [MasterCategories](/javascript/api/outlook/office.mastercategories?view=outlook-js-1.8)</span></span>

<span data-ttu-id="d1b7e-177">このメールボックスのカテゴリマスターリストを管理するためのメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-177">Gets an object that provides methods to manage the categories master list on this mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="d1b7e-178">このメンバーは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-178">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="d1b7e-179">種類</span><span class="sxs-lookup"><span data-stu-id="d1b7e-179">Type</span></span>

*   [<span data-ttu-id="d1b7e-180">MasterCategories</span><span class="sxs-lookup"><span data-stu-id="d1b7e-180">MasterCategories</span></span>](/javascript/api/outlook/office.mastercategories?view=outlook-js-1.8)

##### <a name="requirements"></a><span data-ttu-id="d1b7e-181">Requirements</span><span class="sxs-lookup"><span data-stu-id="d1b7e-181">Requirements</span></span>

|<span data-ttu-id="d1b7e-182">要件</span><span class="sxs-lookup"><span data-stu-id="d1b7e-182">Requirement</span></span>| <span data-ttu-id="d1b7e-183">値</span><span class="sxs-lookup"><span data-stu-id="d1b7e-183">Value</span></span>|
|---|---|
|[<span data-ttu-id="d1b7e-184">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d1b7e-184">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d1b7e-185">1.8</span><span class="sxs-lookup"><span data-stu-id="d1b7e-185">1.8</span></span> |
|[<span data-ttu-id="d1b7e-186">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d1b7e-186">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d1b7e-187">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="d1b7e-187">ReadWriteMailbox</span></span> |
|[<span data-ttu-id="d1b7e-188">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d1b7e-188">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d1b7e-189">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d1b7e-189">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="d1b7e-190">例</span><span class="sxs-lookup"><span data-stu-id="d1b7e-190">Example</span></span>

<span data-ttu-id="d1b7e-191">この例では、このメールボックスのカテゴリマスターリストを取得します。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-191">This example gets the categories master list for this mailbox.</span></span>

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

#### <a name="resturl-string"></a><span data-ttu-id="d1b7e-192">restUrl: String</span><span class="sxs-lookup"><span data-stu-id="d1b7e-192">restUrl: String</span></span>

<span data-ttu-id="d1b7e-193">この電子メール アカウントの REST エンドポイントの URL を取得します。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-193">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="d1b7e-194">`restUrl` 値は、ユーザーのメールボックスに [REST API](/outlook/rest/) 呼び出しを行うために使用できます。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-194">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="d1b7e-195">型</span><span class="sxs-lookup"><span data-stu-id="d1b7e-195">Type</span></span>

*   <span data-ttu-id="d1b7e-196">String</span><span class="sxs-lookup"><span data-stu-id="d1b7e-196">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d1b7e-197">Requirements</span><span class="sxs-lookup"><span data-stu-id="d1b7e-197">Requirements</span></span>

|<span data-ttu-id="d1b7e-198">要件</span><span class="sxs-lookup"><span data-stu-id="d1b7e-198">Requirement</span></span>| <span data-ttu-id="d1b7e-199">値</span><span class="sxs-lookup"><span data-stu-id="d1b7e-199">Value</span></span>|
|---|---|
|[<span data-ttu-id="d1b7e-200">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d1b7e-200">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d1b7e-201">1.5</span><span class="sxs-lookup"><span data-stu-id="d1b7e-201">1.5</span></span> |
|[<span data-ttu-id="d1b7e-202">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d1b7e-202">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d1b7e-203">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d1b7e-203">ReadItem</span></span>|
|[<span data-ttu-id="d1b7e-204">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d1b7e-204">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d1b7e-205">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d1b7e-205">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="d1b7e-206">メソッド</span><span class="sxs-lookup"><span data-stu-id="d1b7e-206">Methods</span></span>

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="d1b7e-207">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="d1b7e-207">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="d1b7e-208">サポートされているイベントのイベント ハンドラーを追加します。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-208">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="d1b7e-209">現在、サポートされている`Office.EventType.ItemChanged`イベント`Office.EventType.OfficeThemeChanged`の種類はおよびです。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-209">Currently, the supported event types are `Office.EventType.ItemChanged` and `Office.EventType.OfficeThemeChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d1b7e-210">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d1b7e-210">Parameters</span></span>

| <span data-ttu-id="d1b7e-211">名前</span><span class="sxs-lookup"><span data-stu-id="d1b7e-211">Name</span></span> | <span data-ttu-id="d1b7e-212">種類</span><span class="sxs-lookup"><span data-stu-id="d1b7e-212">Type</span></span> | <span data-ttu-id="d1b7e-213">属性</span><span class="sxs-lookup"><span data-stu-id="d1b7e-213">Attributes</span></span> | <span data-ttu-id="d1b7e-214">説明</span><span class="sxs-lookup"><span data-stu-id="d1b7e-214">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="d1b7e-215">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="d1b7e-215">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="d1b7e-216">ハンドラーを呼び出す必要のあるイベント。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-216">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="d1b7e-217">Function</span><span class="sxs-lookup"><span data-stu-id="d1b7e-217">Function</span></span> || <span data-ttu-id="d1b7e-p104">イベントを処理する関数。関数は、オブジェクト リテラルである単一パラメーターを受け入れる必要があります。パラメーターの `type` プロパティは、`addHandlerAsync` に渡される `eventType` パラメーターと一致します。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-p104">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="d1b7e-221">Object</span><span class="sxs-lookup"><span data-stu-id="d1b7e-221">Object</span></span> | <span data-ttu-id="d1b7e-222">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="d1b7e-222">&lt;optional&gt;</span></span> | <span data-ttu-id="d1b7e-223">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-223">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="d1b7e-224">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d1b7e-224">Object</span></span> | <span data-ttu-id="d1b7e-225">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="d1b7e-225">&lt;optional&gt;</span></span> | <span data-ttu-id="d1b7e-226">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-226">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="d1b7e-227">function</span><span class="sxs-lookup"><span data-stu-id="d1b7e-227">function</span></span>| <span data-ttu-id="d1b7e-228">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d1b7e-228">&lt;optional&gt;</span></span>|<span data-ttu-id="d1b7e-229">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-229">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d1b7e-230">要件</span><span class="sxs-lookup"><span data-stu-id="d1b7e-230">Requirements</span></span>

|<span data-ttu-id="d1b7e-231">要件</span><span class="sxs-lookup"><span data-stu-id="d1b7e-231">Requirement</span></span>| <span data-ttu-id="d1b7e-232">値</span><span class="sxs-lookup"><span data-stu-id="d1b7e-232">Value</span></span>|
|---|---|
|[<span data-ttu-id="d1b7e-233">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d1b7e-233">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d1b7e-234">1.5</span><span class="sxs-lookup"><span data-stu-id="d1b7e-234">1.5</span></span> |
|[<span data-ttu-id="d1b7e-235">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d1b7e-235">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d1b7e-236">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d1b7e-236">ReadItem</span></span> |
|[<span data-ttu-id="d1b7e-237">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d1b7e-237">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d1b7e-238">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d1b7e-238">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d1b7e-239">例</span><span class="sxs-lookup"><span data-stu-id="d1b7e-239">Example</span></span>

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

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="d1b7e-240">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="d1b7e-240">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="d1b7e-241">REST 形式のアイテム ID を EWS 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-241">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="d1b7e-242">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-242">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="d1b7e-p105">REST API ([Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) や [Microsoft Graph](https://graph.microsoft.io/) など) で取得されたアイテム ID は、Exchange Web サービス (EWS) に使用される形式とは異なる形式を使用します。`convertToEwsId` メソッドは、REST 形式の ID を EWS 用の適切な形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-p105">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d1b7e-245">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d1b7e-245">Parameters</span></span>

|<span data-ttu-id="d1b7e-246">名前</span><span class="sxs-lookup"><span data-stu-id="d1b7e-246">Name</span></span>| <span data-ttu-id="d1b7e-247">型</span><span class="sxs-lookup"><span data-stu-id="d1b7e-247">Type</span></span>| <span data-ttu-id="d1b7e-248">説明</span><span class="sxs-lookup"><span data-stu-id="d1b7e-248">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="d1b7e-249">String</span><span class="sxs-lookup"><span data-stu-id="d1b7e-249">String</span></span>|<span data-ttu-id="d1b7e-250">Outlook REST API 形式のアイテム ID</span><span class="sxs-lookup"><span data-stu-id="d1b7e-250">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="d1b7e-251">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="d1b7e-251">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.8)|<span data-ttu-id="d1b7e-252">アイテム ID の取得に使用された Outlook REST API のバージョンを示す値。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-252">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d1b7e-253">Requirements</span><span class="sxs-lookup"><span data-stu-id="d1b7e-253">Requirements</span></span>

|<span data-ttu-id="d1b7e-254">要件</span><span class="sxs-lookup"><span data-stu-id="d1b7e-254">Requirement</span></span>| <span data-ttu-id="d1b7e-255">値</span><span class="sxs-lookup"><span data-stu-id="d1b7e-255">Value</span></span>|
|---|---|
|[<span data-ttu-id="d1b7e-256">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d1b7e-256">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d1b7e-257">1.3</span><span class="sxs-lookup"><span data-stu-id="d1b7e-257">1.3</span></span>|
|[<span data-ttu-id="d1b7e-258">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d1b7e-258">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d1b7e-259">制限あり</span><span class="sxs-lookup"><span data-stu-id="d1b7e-259">Restricted</span></span>|
|[<span data-ttu-id="d1b7e-260">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d1b7e-260">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d1b7e-261">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d1b7e-261">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d1b7e-262">戻り値:</span><span class="sxs-lookup"><span data-stu-id="d1b7e-262">Returns:</span></span>

<span data-ttu-id="d1b7e-263">型:String</span><span class="sxs-lookup"><span data-stu-id="d1b7e-263">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="d1b7e-264">例</span><span class="sxs-lookup"><span data-stu-id="d1b7e-264">Example</span></span>

```js
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-18"></a><span data-ttu-id="d1b7e-265">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.8)}</span><span class="sxs-lookup"><span data-stu-id="d1b7e-265">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.8)}</span></span>

<span data-ttu-id="d1b7e-266">クライアントのローカル時間で時間情報が含まれている辞書を取得します。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-266">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="d1b7e-p106">Outlook on the web または Outlook デスクトップのメールアプリでは、日付と時刻に異なるタイムゾーンを使用できます。Outlook デスクトップは、クライアント コンピューターのタイム ゾーンを使用します。Outlook on the web は、Exchange 管理センター (EAC) で設定されたタイム ゾーンを使用します。日付と時刻の値は、ユーザー インターフェイスに表示される値が、常にユーザーが期待するタイム ゾーンと一致するように処理する必要があります。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-p106">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times. Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="d1b7e-p107">Outlook デスクトップ クライアントでメール アプリを実行している場合、`convertToLocalClientTime` メソッドは、クライアント コンピューターのタイム ゾーンに設定された値のディクショナリ オブジェクトを返します。Outlook on the web でメール アプリを実行している場合、`convertToLocalClientTime` メソッドは、EAC で指定したタイム ゾーンに設定された値のディクショナリ オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-p107">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d1b7e-272">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d1b7e-272">Parameters</span></span>

|<span data-ttu-id="d1b7e-273">名前</span><span class="sxs-lookup"><span data-stu-id="d1b7e-273">Name</span></span>| <span data-ttu-id="d1b7e-274">型</span><span class="sxs-lookup"><span data-stu-id="d1b7e-274">Type</span></span>| <span data-ttu-id="d1b7e-275">説明</span><span class="sxs-lookup"><span data-stu-id="d1b7e-275">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="d1b7e-276">日付</span><span class="sxs-lookup"><span data-stu-id="d1b7e-276">Date</span></span>|<span data-ttu-id="d1b7e-277">日付オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d1b7e-277">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d1b7e-278">Requirements</span><span class="sxs-lookup"><span data-stu-id="d1b7e-278">Requirements</span></span>

|<span data-ttu-id="d1b7e-279">要件</span><span class="sxs-lookup"><span data-stu-id="d1b7e-279">Requirement</span></span>| <span data-ttu-id="d1b7e-280">値</span><span class="sxs-lookup"><span data-stu-id="d1b7e-280">Value</span></span>|
|---|---|
|[<span data-ttu-id="d1b7e-281">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d1b7e-281">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d1b7e-282">1.0</span><span class="sxs-lookup"><span data-stu-id="d1b7e-282">1.0</span></span>|
|[<span data-ttu-id="d1b7e-283">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d1b7e-283">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d1b7e-284">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d1b7e-284">ReadItem</span></span>|
|[<span data-ttu-id="d1b7e-285">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d1b7e-285">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d1b7e-286">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d1b7e-286">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d1b7e-287">戻り値:</span><span class="sxs-lookup"><span data-stu-id="d1b7e-287">Returns:</span></span>

<span data-ttu-id="d1b7e-288">型:[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="d1b7e-288">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.8)</span></span>

<br>

---
---

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="d1b7e-289">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="d1b7e-289">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="d1b7e-290">EWS 形式のアイテム ID を REST 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-290">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="d1b7e-291">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-291">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="d1b7e-p108">EWS または `itemId` プロパティで取得されるアイテム ID は、REST API ([Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) や [Microsoft Graph](https://graph.microsoft.io/) など) に使用される形式とは異なる形式を使用します。`convertToRestId` メソッドは、EWS 形式の ID を REST 用の適切な形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-p108">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d1b7e-294">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d1b7e-294">Parameters</span></span>

|<span data-ttu-id="d1b7e-295">名前</span><span class="sxs-lookup"><span data-stu-id="d1b7e-295">Name</span></span>| <span data-ttu-id="d1b7e-296">型</span><span class="sxs-lookup"><span data-stu-id="d1b7e-296">Type</span></span>| <span data-ttu-id="d1b7e-297">説明</span><span class="sxs-lookup"><span data-stu-id="d1b7e-297">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="d1b7e-298">String</span><span class="sxs-lookup"><span data-stu-id="d1b7e-298">String</span></span>|<span data-ttu-id="d1b7e-299">Exchange Web サービス (EWS) 形式のアイテム ID</span><span class="sxs-lookup"><span data-stu-id="d1b7e-299">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="d1b7e-300">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="d1b7e-300">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.8)|<span data-ttu-id="d1b7e-301">変換後の ID を使用する Outlook REST API のバージョンを示す値。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-301">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d1b7e-302">Requirements</span><span class="sxs-lookup"><span data-stu-id="d1b7e-302">Requirements</span></span>

|<span data-ttu-id="d1b7e-303">要件</span><span class="sxs-lookup"><span data-stu-id="d1b7e-303">Requirement</span></span>| <span data-ttu-id="d1b7e-304">値</span><span class="sxs-lookup"><span data-stu-id="d1b7e-304">Value</span></span>|
|---|---|
|[<span data-ttu-id="d1b7e-305">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d1b7e-305">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d1b7e-306">1.3</span><span class="sxs-lookup"><span data-stu-id="d1b7e-306">1.3</span></span>|
|[<span data-ttu-id="d1b7e-307">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d1b7e-307">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d1b7e-308">制限あり</span><span class="sxs-lookup"><span data-stu-id="d1b7e-308">Restricted</span></span>|
|[<span data-ttu-id="d1b7e-309">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d1b7e-309">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d1b7e-310">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d1b7e-310">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d1b7e-311">戻り値:</span><span class="sxs-lookup"><span data-stu-id="d1b7e-311">Returns:</span></span>

<span data-ttu-id="d1b7e-312">型:String</span><span class="sxs-lookup"><span data-stu-id="d1b7e-312">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="d1b7e-313">例</span><span class="sxs-lookup"><span data-stu-id="d1b7e-313">Example</span></span>

```js
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="d1b7e-314">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="d1b7e-314">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="d1b7e-315">時間情報が含まれているディクショナリから日付オブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-315">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="d1b7e-316">`convertToUtcClientTime` メソッドは、ローカルの日付と時刻を含むディクショナリを、ローカルの日付と時刻の正しい値を持つ日付オブジェクトに変換します。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-316">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d1b7e-317">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d1b7e-317">Parameters</span></span>

|<span data-ttu-id="d1b7e-318">名前</span><span class="sxs-lookup"><span data-stu-id="d1b7e-318">Name</span></span>| <span data-ttu-id="d1b7e-319">型</span><span class="sxs-lookup"><span data-stu-id="d1b7e-319">Type</span></span>| <span data-ttu-id="d1b7e-320">説明</span><span class="sxs-lookup"><span data-stu-id="d1b7e-320">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="d1b7e-321">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="d1b7e-321">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.8)|<span data-ttu-id="d1b7e-322">変換するローカル時刻の値。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-322">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d1b7e-323">Requirements</span><span class="sxs-lookup"><span data-stu-id="d1b7e-323">Requirements</span></span>

|<span data-ttu-id="d1b7e-324">要件</span><span class="sxs-lookup"><span data-stu-id="d1b7e-324">Requirement</span></span>| <span data-ttu-id="d1b7e-325">値</span><span class="sxs-lookup"><span data-stu-id="d1b7e-325">Value</span></span>|
|---|---|
|[<span data-ttu-id="d1b7e-326">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d1b7e-326">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d1b7e-327">1.0</span><span class="sxs-lookup"><span data-stu-id="d1b7e-327">1.0</span></span>|
|[<span data-ttu-id="d1b7e-328">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d1b7e-328">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d1b7e-329">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d1b7e-329">ReadItem</span></span>|
|[<span data-ttu-id="d1b7e-330">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d1b7e-330">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d1b7e-331">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d1b7e-331">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d1b7e-332">戻り値:</span><span class="sxs-lookup"><span data-stu-id="d1b7e-332">Returns:</span></span>

<span data-ttu-id="d1b7e-333">時間が UTC で表現された日付オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-333">A Date object with the time expressed in UTC.</span></span>

<span data-ttu-id="d1b7e-334">型: Date</span><span class="sxs-lookup"><span data-stu-id="d1b7e-334">Type: Date</span></span>

##### <a name="example"></a><span data-ttu-id="d1b7e-335">例</span><span class="sxs-lookup"><span data-stu-id="d1b7e-335">Example</span></span>

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

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="d1b7e-336">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="d1b7e-336">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="d1b7e-337">既存の予定を表示します。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-337">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="d1b7e-338">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-338">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="d1b7e-339">`displayAppointmentForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存の予定を開きます。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-339">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="d1b7e-p109">Outlook on Mac では、このメソッドを使用して定期的な系列に含まれない単発の予定や定期的な系列のマスター予定を表示できます。ただし、系列のインスタンスは表示できません。これは、Outlook on Mac では定期的な系列のインスタンスのプロパティ (アイテム ID を含む) にアクセスできないためです。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-p109">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="d1b7e-342">Outlook on the web では、このメソッドはフォームの本文が 32KB 以下の文字数の場合にのみ指定のフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-342">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="d1b7e-343">指定のアイテム識別子が既存の予定を表していない場合は、クライアント コンピューターまたはデバイスで空のウィンドウが開き、エラー メッセージは返されません。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-343">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d1b7e-344">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d1b7e-344">Parameters</span></span>

|<span data-ttu-id="d1b7e-345">名前</span><span class="sxs-lookup"><span data-stu-id="d1b7e-345">Name</span></span>| <span data-ttu-id="d1b7e-346">型</span><span class="sxs-lookup"><span data-stu-id="d1b7e-346">Type</span></span>| <span data-ttu-id="d1b7e-347">説明</span><span class="sxs-lookup"><span data-stu-id="d1b7e-347">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="d1b7e-348">String</span><span class="sxs-lookup"><span data-stu-id="d1b7e-348">String</span></span>|<span data-ttu-id="d1b7e-349">既存の予定の Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-349">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d1b7e-350">要件</span><span class="sxs-lookup"><span data-stu-id="d1b7e-350">Requirements</span></span>

|<span data-ttu-id="d1b7e-351">要件</span><span class="sxs-lookup"><span data-stu-id="d1b7e-351">Requirement</span></span>| <span data-ttu-id="d1b7e-352">値</span><span class="sxs-lookup"><span data-stu-id="d1b7e-352">Value</span></span>|
|---|---|
|[<span data-ttu-id="d1b7e-353">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d1b7e-353">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d1b7e-354">1.0</span><span class="sxs-lookup"><span data-stu-id="d1b7e-354">1.0</span></span>|
|[<span data-ttu-id="d1b7e-355">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d1b7e-355">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d1b7e-356">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d1b7e-356">ReadItem</span></span>|
|[<span data-ttu-id="d1b7e-357">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d1b7e-357">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d1b7e-358">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d1b7e-358">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d1b7e-359">例</span><span class="sxs-lookup"><span data-stu-id="d1b7e-359">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

<br>

---
---

#### <a name="displaymessageformitemid"></a><span data-ttu-id="d1b7e-360">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="d1b7e-360">displayMessageForm(itemId)</span></span>

<span data-ttu-id="d1b7e-361">既存のメッセージを表示します。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-361">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="d1b7e-362">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-362">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="d1b7e-363">`displayMessageForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存のメッセージを開きます。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-363">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="d1b7e-364">Outlook on the web では、このメソッドはフォームの本文が 32 KB 以下の文字数の場合にのみ指定のフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-364">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="d1b7e-365">指定のアイテム識別子が既存のメッセージを表していない場合は、クライアント コンピューターにはメッセージは表示されず、エラー メッセージも返されません。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-365">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="d1b7e-p110">予定を表す `itemId` を含む `displayMessageForm` を使用しないでください。既存の予定を表示するには、`displayAppointmentForm` メソッドを使用します。新しい予定を作成するフォームを表示するには、`displayNewAppointmentForm` メソッドを使用します。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-p110">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d1b7e-368">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d1b7e-368">Parameters</span></span>

|<span data-ttu-id="d1b7e-369">名前</span><span class="sxs-lookup"><span data-stu-id="d1b7e-369">Name</span></span>| <span data-ttu-id="d1b7e-370">型</span><span class="sxs-lookup"><span data-stu-id="d1b7e-370">Type</span></span>| <span data-ttu-id="d1b7e-371">説明</span><span class="sxs-lookup"><span data-stu-id="d1b7e-371">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="d1b7e-372">String</span><span class="sxs-lookup"><span data-stu-id="d1b7e-372">String</span></span>|<span data-ttu-id="d1b7e-373">既存のメッセージの Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-373">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d1b7e-374">Requirements</span><span class="sxs-lookup"><span data-stu-id="d1b7e-374">Requirements</span></span>

|<span data-ttu-id="d1b7e-375">要件</span><span class="sxs-lookup"><span data-stu-id="d1b7e-375">Requirement</span></span>| <span data-ttu-id="d1b7e-376">値</span><span class="sxs-lookup"><span data-stu-id="d1b7e-376">Value</span></span>|
|---|---|
|[<span data-ttu-id="d1b7e-377">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d1b7e-377">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d1b7e-378">1.0</span><span class="sxs-lookup"><span data-stu-id="d1b7e-378">1.0</span></span>|
|[<span data-ttu-id="d1b7e-379">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d1b7e-379">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d1b7e-380">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d1b7e-380">ReadItem</span></span>|
|[<span data-ttu-id="d1b7e-381">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d1b7e-381">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d1b7e-382">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d1b7e-382">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d1b7e-383">例</span><span class="sxs-lookup"><span data-stu-id="d1b7e-383">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

<br>

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="d1b7e-384">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="d1b7e-384">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="d1b7e-385">新しい予定を作成するためのフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-385">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="d1b7e-386">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-386">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="d1b7e-p111">`displayNewAppointmentForm` メソッドを使用すると、ユーザーが新しい予定または会議を作成できるフォームが開きます。パラメーターを指定すると、予定のフォーム フィールドにパラメーターの内容が自動的に設定されます。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-p111">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="d1b7e-p112">Outlook on the web およびモバイル デバイスでは、このメソッドは常に出席者フィールドが含まれるフォームを表示します。入力引数として出席者を指定しないと、このメソッドは **[保存]** ボタンのあるフォームを表示します。出席者を指定した場合には、フォームにその出席者と **[送信]** ボタンが表示されます。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-p112">In Outlook on the web and mobile devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="d1b7e-p113">Outlook リッチ クライアントと Outlook RT で、`requiredAttendees`、`optionalAttendees`、または `resources` パラメーターに出席者またはリソースを指定し、このメソッドを実行すると、**[送信]** ボタンがある会議フォームが表示されます。受信者を指定せずにこのメソッドを実行すると、**[保存して閉じる]** ボタンがある予定フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-p113">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="d1b7e-394">パラメーターのいずれかが指定のサイズ制限を超える場合、または不明なパラメーター名が指定されている場合は、例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-394">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d1b7e-395">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d1b7e-395">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="d1b7e-396">すべてのパラメーターは省略可能です。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-396">All parameters are optional.</span></span>

|<span data-ttu-id="d1b7e-397">名前</span><span class="sxs-lookup"><span data-stu-id="d1b7e-397">Name</span></span>| <span data-ttu-id="d1b7e-398">型</span><span class="sxs-lookup"><span data-stu-id="d1b7e-398">Type</span></span>| <span data-ttu-id="d1b7e-399">説明</span><span class="sxs-lookup"><span data-stu-id="d1b7e-399">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="d1b7e-400">Object</span><span class="sxs-lookup"><span data-stu-id="d1b7e-400">Object</span></span> | <span data-ttu-id="d1b7e-401">新しい予定を記述するパラメーターのディクショナリ。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-401">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="d1b7e-402">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)&gt;</span><span class="sxs-lookup"><span data-stu-id="d1b7e-402">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)&gt;</span></span> | <span data-ttu-id="d1b7e-p114">予定に必要な各出席者について、メール アドレスを含む文字列の配列、または `EmailAddressDetails` オブジェクトを含む配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-p114">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="d1b7e-405">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)&gt;</span><span class="sxs-lookup"><span data-stu-id="d1b7e-405">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)&gt;</span></span> | <span data-ttu-id="d1b7e-p115">予定の各任意出席者について、メール アドレスを含む文字列の配列、または `EmailAddressDetails` オブジェクトを含む配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-p115">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="d1b7e-408">日付</span><span class="sxs-lookup"><span data-stu-id="d1b7e-408">Date</span></span> | <span data-ttu-id="d1b7e-409">予定の開始日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-409">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="d1b7e-410">日付</span><span class="sxs-lookup"><span data-stu-id="d1b7e-410">Date</span></span> | <span data-ttu-id="d1b7e-411">予定の終了日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-411">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="d1b7e-412">String</span><span class="sxs-lookup"><span data-stu-id="d1b7e-412">String</span></span> | <span data-ttu-id="d1b7e-p116">予定の場所を含む文字列。文字列は最大 255 文字に制限されます。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-p116">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="d1b7e-415">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="d1b7e-415">Array.&lt;String&gt;</span></span> | <span data-ttu-id="d1b7e-p117">予定に必要なリソースを含む文字列の配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-p117">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="d1b7e-418">String</span><span class="sxs-lookup"><span data-stu-id="d1b7e-418">String</span></span> | <span data-ttu-id="d1b7e-p118">予定の件名を含む文字列です。文字列は最大 255 文字に制限されます。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-p118">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="d1b7e-421">String</span><span class="sxs-lookup"><span data-stu-id="d1b7e-421">String</span></span> | <span data-ttu-id="d1b7e-p119">予定の本文。本文の内容は、最大サイズが 32 KB に制限されます。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-p119">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="d1b7e-424">要件</span><span class="sxs-lookup"><span data-stu-id="d1b7e-424">Requirements</span></span>

|<span data-ttu-id="d1b7e-425">要件</span><span class="sxs-lookup"><span data-stu-id="d1b7e-425">Requirement</span></span>| <span data-ttu-id="d1b7e-426">値</span><span class="sxs-lookup"><span data-stu-id="d1b7e-426">Value</span></span>|
|---|---|
|[<span data-ttu-id="d1b7e-427">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d1b7e-427">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d1b7e-428">1.0</span><span class="sxs-lookup"><span data-stu-id="d1b7e-428">1.0</span></span>|
|[<span data-ttu-id="d1b7e-429">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d1b7e-429">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d1b7e-430">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d1b7e-430">ReadItem</span></span>|
|[<span data-ttu-id="d1b7e-431">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d1b7e-431">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d1b7e-432">読み取り</span><span class="sxs-lookup"><span data-stu-id="d1b7e-432">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d1b7e-433">例</span><span class="sxs-lookup"><span data-stu-id="d1b7e-433">Example</span></span>

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

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="d1b7e-434">displayNewMessageForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="d1b7e-434">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="d1b7e-435">新しいメッセージを作成するためのフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-435">Displays a form for creating a new message.</span></span>

<span data-ttu-id="d1b7e-p120">The `displayNewMessageForm` method opens a form that enables the user to create a new message. If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span><span class="sxs-lookup"><span data-stu-id="d1b7e-p120">The `displayNewMessageForm` method opens a form that enables the user to create a new message. If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="d1b7e-438">パラメーターのいずれかが指定のサイズ制限を超える場合、または不明なパラメーター名が指定されている場合は、例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-438">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d1b7e-439">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d1b7e-439">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="d1b7e-440">すべてのパラメーターは省略可能です。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-440">All parameters are optional.</span></span>

|<span data-ttu-id="d1b7e-441">名前</span><span class="sxs-lookup"><span data-stu-id="d1b7e-441">Name</span></span>| <span data-ttu-id="d1b7e-442">型</span><span class="sxs-lookup"><span data-stu-id="d1b7e-442">Type</span></span>| <span data-ttu-id="d1b7e-443">説明</span><span class="sxs-lookup"><span data-stu-id="d1b7e-443">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="d1b7e-444">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d1b7e-444">Object</span></span> | <span data-ttu-id="d1b7e-445">新しいメッセージを記述するパラメータの辞書。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-445">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="d1b7e-446">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)&gt;</span><span class="sxs-lookup"><span data-stu-id="d1b7e-446">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)&gt;</span></span> | <span data-ttu-id="d1b7e-p121">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line. The array is limited to a maximum of 100 entries.</span><span class="sxs-lookup"><span data-stu-id="d1b7e-p121">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="d1b7e-449">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)&gt;</span><span class="sxs-lookup"><span data-stu-id="d1b7e-449">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)&gt;</span></span> | <span data-ttu-id="d1b7e-p122">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line. The array is limited to a maximum of 100 entries.</span><span class="sxs-lookup"><span data-stu-id="d1b7e-p122">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="d1b7e-452">配列。&lt;文字列&gt; | 配列。&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)&gt;</span><span class="sxs-lookup"><span data-stu-id="d1b7e-452">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)&gt;</span></span> | <span data-ttu-id="d1b7e-p123">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line. The array is limited to a maximum of 100 entries.</span><span class="sxs-lookup"><span data-stu-id="d1b7e-p123">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="d1b7e-455">String</span><span class="sxs-lookup"><span data-stu-id="d1b7e-455">String</span></span> | <span data-ttu-id="d1b7e-p124">A string containing the subject of the message. The string is limited to a maximum of 255 characters.</span><span class="sxs-lookup"><span data-stu-id="d1b7e-p124">A string containing the subject of the message. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="d1b7e-458">String</span><span class="sxs-lookup"><span data-stu-id="d1b7e-458">String</span></span> | <span data-ttu-id="d1b7e-p125">The HTML body of the message. The body content is limited to a maximum size of 32 KB.</span><span class="sxs-lookup"><span data-stu-id="d1b7e-p125">The HTML body of the message. The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="d1b7e-461">配列。&lt;オブジェクト&gt;</span><span class="sxs-lookup"><span data-stu-id="d1b7e-461">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="d1b7e-462">添付ファイルまたは添付アイテムである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-462">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="d1b7e-463">String</span><span class="sxs-lookup"><span data-stu-id="d1b7e-463">String</span></span> | <span data-ttu-id="d1b7e-p126">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-p126">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="d1b7e-466">String</span><span class="sxs-lookup"><span data-stu-id="d1b7e-466">String</span></span> | <span data-ttu-id="d1b7e-467">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-467">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="d1b7e-468">文字列</span><span class="sxs-lookup"><span data-stu-id="d1b7e-468">String</span></span> | <span data-ttu-id="d1b7e-p127">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-p127">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="d1b7e-471">ブール値</span><span class="sxs-lookup"><span data-stu-id="d1b7e-471">Boolean</span></span> | <span data-ttu-id="d1b7e-p128">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-p128">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="d1b7e-474">String</span><span class="sxs-lookup"><span data-stu-id="d1b7e-474">String</span></span> | <span data-ttu-id="d1b7e-p129">Only used if `type` is set to `item`. The EWS item id of the existing e-mail you want to attach to the new message. This is a string up to 100 characters.</span><span class="sxs-lookup"><span data-stu-id="d1b7e-p129">Only used if `type` is set to `item`. The EWS item id of the existing e-mail you want to attach to the new message. This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="d1b7e-478">Requirements</span><span class="sxs-lookup"><span data-stu-id="d1b7e-478">Requirements</span></span>

|<span data-ttu-id="d1b7e-479">要件</span><span class="sxs-lookup"><span data-stu-id="d1b7e-479">Requirement</span></span>| <span data-ttu-id="d1b7e-480">値</span><span class="sxs-lookup"><span data-stu-id="d1b7e-480">Value</span></span>|
|---|---|
|[<span data-ttu-id="d1b7e-481">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d1b7e-481">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d1b7e-482">1.6</span><span class="sxs-lookup"><span data-stu-id="d1b7e-482">1.6</span></span> |
|[<span data-ttu-id="d1b7e-483">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d1b7e-483">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d1b7e-484">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d1b7e-484">ReadItem</span></span>|
|[<span data-ttu-id="d1b7e-485">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d1b7e-485">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d1b7e-486">読み取り</span><span class="sxs-lookup"><span data-stu-id="d1b7e-486">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d1b7e-487">例</span><span class="sxs-lookup"><span data-stu-id="d1b7e-487">Example</span></span>

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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="d1b7e-488">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="d1b7e-488">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="d1b7e-489">REST API または Exchange Web サービスを呼び出すために使用するトークンを含む文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-489">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="d1b7e-p130">`getCallbackTokenAsync` メソッドは、ユーザーのメールボックスをホストする Exchange Server から不透明なトークンを取得する非同期の呼び出しを行います。コールバック トークンの有効期間は 5 分です。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-p130">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="d1b7e-492">可能であれば、アドインでは Exchange Web サービスの代わりに REST API を使用することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-492">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span>

<span data-ttu-id="d1b7e-493">閲覧モードで `getCallbackTokenAsync` メソッドを呼び出すには、**ReadItem** の最小限のアクセス許可レベルが必要です。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-493">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="d1b7e-494">作成モードで `getCallbackTokenAsync` を呼び出すには、アイテムを保存しておく必要があります。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-494">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="d1b7e-495">[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドは、**ReadWriteItem** の最小限のアクセス許可レベルが必要です。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-495">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

<span data-ttu-id="d1b7e-496">**REST トークン**</span><span class="sxs-lookup"><span data-stu-id="d1b7e-496">**REST Tokens**</span></span>

<span data-ttu-id="d1b7e-p132">REST トークンが要求された場合 (`options.isRest = true`)、結果トークンは Exchange Web サービスの呼び出しを認証するためには機能しません。アドインがマニフェストで [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) アクセス許可を指定していない限り、トークンの範囲は現在のアイテムとその添付ファイルへの読み取り専用アクセスに制限されます。`ReadWriteMailbox` アクセス許可が指定されている場合は、結果トークンは、メールを送信する機能など、メール、カレンダー、連絡先への読み取り/書き込みアクセスを付与します。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-p132">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="d1b7e-500">アドインでは、`restUrl` プロパティを使用して、REST API 呼び出しを行うときに使用する正しい URL を決定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-500">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="d1b7e-501">**EWS トークン**</span><span class="sxs-lookup"><span data-stu-id="d1b7e-501">**EWS Tokens**</span></span>

<span data-ttu-id="d1b7e-p133">EWS トークンが要求された場合 (`options.isRest = false`)、結果トークンは REST API 呼び出しを認証するためには機能しません。トークンの範囲は、現在のアイテムへのアクセスに制限されます。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-p133">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="d1b7e-504">アドインでは、`ewsUrl` プロパティを使用して、EWS 呼び出しを行うときに使用する正しい URL を決定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-504">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

<span data-ttu-id="d1b7e-505">トークンと、添付ファイル識別子またはアイテム識別子の両方をサードパーティ システムに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-505">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="d1b7e-506">サードパーティシステムは、トークンをベアラー認証トークンとして使用して、Exchange Web サービス (EWS) の[Getattachment](/exchange/client-developer/web-service-reference/getattachment-operation)操作または[GetItem](/exchange/client-developer/web-service-reference/getitem-operation)操作を呼び出して、添付ファイルまたはアイテムを取得します。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-506">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to retrieve an attachment or item.</span></span> <span data-ttu-id="d1b7e-507">たとえば、[選択したアイテムから添付ファイルを取得する](/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-507">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

##### <a name="parameters"></a><span data-ttu-id="d1b7e-508">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d1b7e-508">Parameters</span></span>

|<span data-ttu-id="d1b7e-509">名前</span><span class="sxs-lookup"><span data-stu-id="d1b7e-509">Name</span></span>| <span data-ttu-id="d1b7e-510">型</span><span class="sxs-lookup"><span data-stu-id="d1b7e-510">Type</span></span>| <span data-ttu-id="d1b7e-511">属性</span><span class="sxs-lookup"><span data-stu-id="d1b7e-511">Attributes</span></span>| <span data-ttu-id="d1b7e-512">説明</span><span class="sxs-lookup"><span data-stu-id="d1b7e-512">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="d1b7e-513">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d1b7e-513">Object</span></span> | <span data-ttu-id="d1b7e-514">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="d1b7e-514">&lt;optional&gt;</span></span> | <span data-ttu-id="d1b7e-515">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-515">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="d1b7e-516">ブール値</span><span class="sxs-lookup"><span data-stu-id="d1b7e-516">Boolean</span></span> |  <span data-ttu-id="d1b7e-517">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d1b7e-517">&lt;optional&gt;</span></span> | <span data-ttu-id="d1b7e-p135">提供されたトークンを Outlook REST API または Exchange Web サービスに使用するかどうかを決定します。既定値は、`false` です。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-p135">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="d1b7e-520">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d1b7e-520">Object</span></span> |  <span data-ttu-id="d1b7e-521">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="d1b7e-521">&lt;optional&gt;</span></span> | <span data-ttu-id="d1b7e-522">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-522">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="d1b7e-523">function</span><span class="sxs-lookup"><span data-stu-id="d1b7e-523">function</span></span>||<span data-ttu-id="d1b7e-524">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-524">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="d1b7e-525">トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-525">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="d1b7e-526">エラーが発生した場合、 `asyncResult.error` および `asyncResult.diagnostics` のプロパティで追加情報が提供される場合があります。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-526">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|

##### <a name="errors"></a><span data-ttu-id="d1b7e-527">エラー</span><span class="sxs-lookup"><span data-stu-id="d1b7e-527">Errors</span></span>

|<span data-ttu-id="d1b7e-528">エラー コード</span><span class="sxs-lookup"><span data-stu-id="d1b7e-528">Error code</span></span>|<span data-ttu-id="d1b7e-529">説明</span><span class="sxs-lookup"><span data-stu-id="d1b7e-529">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="d1b7e-530">要求が失敗しました。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-530">The request has failed.</span></span> <span data-ttu-id="d1b7e-531">HTTP エラーコードの diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-531">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="d1b7e-532">Exchange サーバーがエラーを返しました。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-532">The Exchange server returned an error.</span></span> <span data-ttu-id="d1b7e-533">詳細については、diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-533">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="d1b7e-534">ユーザーはネットワークに接続されていません。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-534">The user is no longer connected to the network.</span></span> <span data-ttu-id="d1b7e-535">ネットワーク接続を確認し、やり直してください。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-535">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d1b7e-536">要件</span><span class="sxs-lookup"><span data-stu-id="d1b7e-536">Requirements</span></span>

|<span data-ttu-id="d1b7e-537">必要条件</span><span class="sxs-lookup"><span data-stu-id="d1b7e-537">Requirement</span></span>| <span data-ttu-id="d1b7e-538">値</span><span class="sxs-lookup"><span data-stu-id="d1b7e-538">Value</span></span>|
|---|---|
|[<span data-ttu-id="d1b7e-539">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d1b7e-539">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d1b7e-540">1.5</span><span class="sxs-lookup"><span data-stu-id="d1b7e-540">1.5</span></span> |
|[<span data-ttu-id="d1b7e-541">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d1b7e-541">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d1b7e-542">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d1b7e-542">ReadItem</span></span>|
|[<span data-ttu-id="d1b7e-543">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d1b7e-543">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d1b7e-544">新規作成と閲覧</span><span class="sxs-lookup"><span data-stu-id="d1b7e-544">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="d1b7e-545">例</span><span class="sxs-lookup"><span data-stu-id="d1b7e-545">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="d1b7e-546">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="d1b7e-546">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="d1b7e-547">Exchange Server から添付ファイルやアイテムを取得するために使うトークンを含む文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-547">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="d1b7e-p139">`getCallbackTokenAsync` メソッドは、ユーザーのメールボックスをホストする Exchange Server から不透明なトークンを取得する非同期の呼び出しを行います。コールバック トークンの有効期間は 5 分です。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-p139">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="d1b7e-550">トークンと、添付ファイル識別子またはアイテム識別子の両方をサードパーティ システムに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-550">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="d1b7e-551">サードパーティ システムは、トークンをベアラー承認トークンとして使用し、Exchange Web サービス (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) 操作または [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) 操作を呼び出して、添付ファイルまたはアイテムを返します。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-551">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item.</span></span> <span data-ttu-id="d1b7e-552">たとえば、[選択したアイテムから添付ファイルを取得する](/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-552">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="d1b7e-553">閲覧モードで `getCallbackTokenAsync` メソッドを呼び出すには、**ReadItem** の最小限のアクセス許可レベルが必要です。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-553">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="d1b7e-554">作成モードで `getCallbackTokenAsync` を呼び出すには、アイテムを保存しておく必要があります。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-554">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="d1b7e-555">[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドは、**ReadWriteItem** の最小限のアクセス許可レベルが必要です。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-555">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d1b7e-556">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d1b7e-556">Parameters</span></span>

|<span data-ttu-id="d1b7e-557">名前</span><span class="sxs-lookup"><span data-stu-id="d1b7e-557">Name</span></span>| <span data-ttu-id="d1b7e-558">型</span><span class="sxs-lookup"><span data-stu-id="d1b7e-558">Type</span></span>| <span data-ttu-id="d1b7e-559">属性</span><span class="sxs-lookup"><span data-stu-id="d1b7e-559">Attributes</span></span>| <span data-ttu-id="d1b7e-560">説明</span><span class="sxs-lookup"><span data-stu-id="d1b7e-560">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="d1b7e-561">関数</span><span class="sxs-lookup"><span data-stu-id="d1b7e-561">function</span></span>||<span data-ttu-id="d1b7e-562">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-562">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="d1b7e-563">トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-563">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="d1b7e-564">エラーが発生した場合、 `asyncResult.error` および `asyncResult.diagnostics` のプロパティで追加情報が提供される場合があります。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-564">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="d1b7e-565">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d1b7e-565">Object</span></span>| <span data-ttu-id="d1b7e-566">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="d1b7e-566">&lt;optional&gt;</span></span>|<span data-ttu-id="d1b7e-567">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-567">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="d1b7e-568">エラー</span><span class="sxs-lookup"><span data-stu-id="d1b7e-568">Errors</span></span>

|<span data-ttu-id="d1b7e-569">エラー コード</span><span class="sxs-lookup"><span data-stu-id="d1b7e-569">Error code</span></span>|<span data-ttu-id="d1b7e-570">説明</span><span class="sxs-lookup"><span data-stu-id="d1b7e-570">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="d1b7e-571">要求が失敗しました。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-571">The request has failed.</span></span> <span data-ttu-id="d1b7e-572">HTTP エラーコードの diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-572">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="d1b7e-573">Exchange サーバーがエラーを返しました。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-573">The Exchange server returned an error.</span></span> <span data-ttu-id="d1b7e-574">詳細については、diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-574">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="d1b7e-575">ユーザーはネットワークに接続されていません。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-575">The user is no longer connected to the network.</span></span> <span data-ttu-id="d1b7e-576">ネットワーク接続を確認し、やり直してください。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-576">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d1b7e-577">要件</span><span class="sxs-lookup"><span data-stu-id="d1b7e-577">Requirements</span></span>

|<span data-ttu-id="d1b7e-578">要件</span><span class="sxs-lookup"><span data-stu-id="d1b7e-578">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="d1b7e-579">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d1b7e-579">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d1b7e-580">1.0</span><span class="sxs-lookup"><span data-stu-id="d1b7e-580">1.0</span></span> | <span data-ttu-id="d1b7e-581">1.3</span><span class="sxs-lookup"><span data-stu-id="d1b7e-581">1.3</span></span> |
|[<span data-ttu-id="d1b7e-582">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d1b7e-582">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d1b7e-583">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d1b7e-583">ReadItem</span></span> | <span data-ttu-id="d1b7e-584">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d1b7e-584">ReadItem</span></span> |
|[<span data-ttu-id="d1b7e-585">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d1b7e-585">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d1b7e-586">Read</span><span class="sxs-lookup"><span data-stu-id="d1b7e-586">Read</span></span> | <span data-ttu-id="d1b7e-587">Compose</span><span class="sxs-lookup"><span data-stu-id="d1b7e-587">Compose</span></span> |

##### <a name="example"></a><span data-ttu-id="d1b7e-588">例</span><span class="sxs-lookup"><span data-stu-id="d1b7e-588">Example</span></span>

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

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="d1b7e-589">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="d1b7e-589">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="d1b7e-590">ユーザーと Office アドインを識別するトークンを取得します。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-590">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="d1b7e-591">`getUserIdentityTokenAsync` メソッドは、[アドインとユーザーをサード パーティのシステムで識別して認証](/outlook/add-ins/authentication)することのできるトークンを返します。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-591">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="d1b7e-592">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d1b7e-592">Parameters</span></span>

|<span data-ttu-id="d1b7e-593">名前</span><span class="sxs-lookup"><span data-stu-id="d1b7e-593">Name</span></span>| <span data-ttu-id="d1b7e-594">型</span><span class="sxs-lookup"><span data-stu-id="d1b7e-594">Type</span></span>| <span data-ttu-id="d1b7e-595">属性</span><span class="sxs-lookup"><span data-stu-id="d1b7e-595">Attributes</span></span>| <span data-ttu-id="d1b7e-596">説明</span><span class="sxs-lookup"><span data-stu-id="d1b7e-596">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="d1b7e-597">関数</span><span class="sxs-lookup"><span data-stu-id="d1b7e-597">function</span></span>||<span data-ttu-id="d1b7e-598">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-598">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="d1b7e-599">トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-599">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="d1b7e-600">エラーが発生した場合、 `asyncResult.error` および `asyncResult.diagnostics` のプロパティで追加情報が提供される場合があります。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-600">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="d1b7e-601">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d1b7e-601">Object</span></span>| <span data-ttu-id="d1b7e-602">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="d1b7e-602">&lt;optional&gt;</span></span>|<span data-ttu-id="d1b7e-603">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-603">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="d1b7e-604">エラー</span><span class="sxs-lookup"><span data-stu-id="d1b7e-604">Errors</span></span>

|<span data-ttu-id="d1b7e-605">エラー コード</span><span class="sxs-lookup"><span data-stu-id="d1b7e-605">Error code</span></span>|<span data-ttu-id="d1b7e-606">説明</span><span class="sxs-lookup"><span data-stu-id="d1b7e-606">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="d1b7e-607">要求が失敗しました。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-607">The request has failed.</span></span> <span data-ttu-id="d1b7e-608">HTTP エラーコードの diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-608">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="d1b7e-609">Exchange サーバーがエラーを返しました。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-609">The Exchange server returned an error.</span></span> <span data-ttu-id="d1b7e-610">詳細については、diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-610">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="d1b7e-611">ユーザーはネットワークに接続されていません。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-611">The user is no longer connected to the network.</span></span> <span data-ttu-id="d1b7e-612">ネットワーク接続を確認し、やり直してください。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-612">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d1b7e-613">要件</span><span class="sxs-lookup"><span data-stu-id="d1b7e-613">Requirements</span></span>

|<span data-ttu-id="d1b7e-614">要件</span><span class="sxs-lookup"><span data-stu-id="d1b7e-614">Requirement</span></span>| <span data-ttu-id="d1b7e-615">値</span><span class="sxs-lookup"><span data-stu-id="d1b7e-615">Value</span></span>|
|---|---|
|[<span data-ttu-id="d1b7e-616">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d1b7e-616">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d1b7e-617">1.0</span><span class="sxs-lookup"><span data-stu-id="d1b7e-617">1.0</span></span>|
|[<span data-ttu-id="d1b7e-618">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d1b7e-618">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d1b7e-619">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d1b7e-619">ReadItem</span></span>|
|[<span data-ttu-id="d1b7e-620">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d1b7e-620">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d1b7e-621">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d1b7e-621">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d1b7e-622">例</span><span class="sxs-lookup"><span data-stu-id="d1b7e-622">Example</span></span>

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

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="d1b7e-623">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="d1b7e-623">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="d1b7e-624">ユーザーのメールボックスをホストしている Exchange サーバー上の Exchange Web サービス (EWS) のサービスに対して非同期の要求を行います。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-624">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="d1b7e-625">このメソッドは、次のシナリオではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-625">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="d1b7e-626">Outlook on iOS または Android の場合</span><span class="sxs-lookup"><span data-stu-id="d1b7e-626">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="d1b7e-627">アドインが Gmail のメールボックスに読み込まれる場合</span><span class="sxs-lookup"><span data-stu-id="d1b7e-627">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="d1b7e-628">このような場合は、アドインでは [REST API を使用](/outlook/add-ins/use-rest-api)して、代わりにユーザーのメールボックスにアクセスする必要があります。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-628">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="d1b7e-p148">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange. See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span><span class="sxs-lookup"><span data-stu-id="d1b7e-p148">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange. See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="d1b7e-631">`makeEwsRequestAsync` メソッドでは、フォルダー関連アイテムを要求できません。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-631">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="d1b7e-632">XML 要求では UTF-8 エンコードを指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-632">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="d1b7e-p149">`makeEwsRequestAsync` メソッドを使用するには、アドインに **ReadWriteMailbox** アクセス許可が必要です。**ReadWriteMailbox** アクセス許可と、`makeEwsRequestAsync` メソッドで呼び出せる EWS 操作の使い方については、「[ユーザーのメールボックスへのメール アドイン アクセスのアクセス許可を指定する](/outlook/add-ins/understanding-outlook-add-in-permissions)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-p149">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="d1b7e-635">サーバー管理者は、クライアント アクセス サーバーの EWS ディレクトリで `OAuthAuthentication` を true に設定して、`makeEwsRequestAsync` メソッドで EWS 要求を行うことができるようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-635">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="d1b7e-636">バージョンの相違点</span><span class="sxs-lookup"><span data-stu-id="d1b7e-636">Version differences</span></span>

<span data-ttu-id="d1b7e-637">バージョン 15.0.4535.1004 より前のバージョンの Outlook で実行しているメール アプリで `makeEwsRequestAsync` メソッドを使用する場合には、エンコード値を `ISO-8859-1` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-637">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="d1b7e-638">Outlook on the web でメール アプリを実行している場合は、エンコード値を設定する必要はありません。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-638">You do not need to set the encoding value when your mail app is running in Outlook on the web.</span></span> <span data-ttu-id="d1b7e-639">メールアプリが web 上の Outlook またはデスクトップクライアントで実行されているかどうかは、mailbox プロパティを使用して判断できます。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-639">You can determine whether your mail app is running in Outlook on the web or a desktop client by using the mailbox.diagnostics.hostName property.</span></span> <span data-ttu-id="d1b7e-640">mailbox.diagnostics.hostVersion プロパティを使って、どのバージョンの Outlook を使って実行しているかを確認できます。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-640">You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d1b7e-641">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d1b7e-641">Parameters</span></span>

|<span data-ttu-id="d1b7e-642">名前</span><span class="sxs-lookup"><span data-stu-id="d1b7e-642">Name</span></span>| <span data-ttu-id="d1b7e-643">型</span><span class="sxs-lookup"><span data-stu-id="d1b7e-643">Type</span></span>| <span data-ttu-id="d1b7e-644">属性</span><span class="sxs-lookup"><span data-stu-id="d1b7e-644">Attributes</span></span>| <span data-ttu-id="d1b7e-645">説明</span><span class="sxs-lookup"><span data-stu-id="d1b7e-645">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="d1b7e-646">String</span><span class="sxs-lookup"><span data-stu-id="d1b7e-646">String</span></span>||<span data-ttu-id="d1b7e-647">EWS 要求です。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-647">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="d1b7e-648">function</span><span class="sxs-lookup"><span data-stu-id="d1b7e-648">function</span></span>||<span data-ttu-id="d1b7e-649">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-649">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="d1b7e-p151">The XML result of the EWS call is provided as a string in the `asyncResult.value` property. If the result exceeds 1 MB in size, an error message is returned instead.</span><span class="sxs-lookup"><span data-stu-id="d1b7e-p151">The XML result of the EWS call is provided as a string in the `asyncResult.value` property. If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="d1b7e-652">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d1b7e-652">Object</span></span>| <span data-ttu-id="d1b7e-653">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="d1b7e-653">&lt;optional&gt;</span></span>|<span data-ttu-id="d1b7e-654">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-654">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d1b7e-655">要件</span><span class="sxs-lookup"><span data-stu-id="d1b7e-655">Requirements</span></span>

|<span data-ttu-id="d1b7e-656">要件</span><span class="sxs-lookup"><span data-stu-id="d1b7e-656">Requirement</span></span>| <span data-ttu-id="d1b7e-657">値</span><span class="sxs-lookup"><span data-stu-id="d1b7e-657">Value</span></span>|
|---|---|
|[<span data-ttu-id="d1b7e-658">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d1b7e-658">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d1b7e-659">1.0</span><span class="sxs-lookup"><span data-stu-id="d1b7e-659">1.0</span></span>|
|[<span data-ttu-id="d1b7e-660">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d1b7e-660">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d1b7e-661">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="d1b7e-661">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="d1b7e-662">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d1b7e-662">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d1b7e-663">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d1b7e-663">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d1b7e-664">例</span><span class="sxs-lookup"><span data-stu-id="d1b7e-664">Example</span></span>

<span data-ttu-id="d1b7e-665">次の例は、`GetItem` 操作を使ってアイテムの件名を取得するため、`makeEwsRequestAsync` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-665">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="d1b7e-666">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="d1b7e-666">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="d1b7e-667">サポートされているイベントの種類のイベント ハンドラーを削除します。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-667">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="d1b7e-668">現在、サポートされている`Office.EventType.ItemChanged`イベント`Office.EventType.OfficeThemeChanged`の種類はおよびです。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-668">Currently, the supported event types are `Office.EventType.ItemChanged` and `Office.EventType.OfficeThemeChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d1b7e-669">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d1b7e-669">Parameters</span></span>

| <span data-ttu-id="d1b7e-670">名前</span><span class="sxs-lookup"><span data-stu-id="d1b7e-670">Name</span></span> | <span data-ttu-id="d1b7e-671">型</span><span class="sxs-lookup"><span data-stu-id="d1b7e-671">Type</span></span> | <span data-ttu-id="d1b7e-672">属性</span><span class="sxs-lookup"><span data-stu-id="d1b7e-672">Attributes</span></span> | <span data-ttu-id="d1b7e-673">説明</span><span class="sxs-lookup"><span data-stu-id="d1b7e-673">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="d1b7e-674">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="d1b7e-674">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="d1b7e-675">ハンドラーを取り消すイベント。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-675">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="d1b7e-676">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d1b7e-676">Object</span></span> | <span data-ttu-id="d1b7e-677">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="d1b7e-677">&lt;optional&gt;</span></span> | <span data-ttu-id="d1b7e-678">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-678">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="d1b7e-679">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d1b7e-679">Object</span></span> | <span data-ttu-id="d1b7e-680">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="d1b7e-680">&lt;optional&gt;</span></span> | <span data-ttu-id="d1b7e-681">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-681">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="d1b7e-682">function</span><span class="sxs-lookup"><span data-stu-id="d1b7e-682">function</span></span>| <span data-ttu-id="d1b7e-683">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d1b7e-683">&lt;optional&gt;</span></span>|<span data-ttu-id="d1b7e-684">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="d1b7e-684">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d1b7e-685">Requirements</span><span class="sxs-lookup"><span data-stu-id="d1b7e-685">Requirements</span></span>

|<span data-ttu-id="d1b7e-686">要件</span><span class="sxs-lookup"><span data-stu-id="d1b7e-686">Requirement</span></span>| <span data-ttu-id="d1b7e-687">値</span><span class="sxs-lookup"><span data-stu-id="d1b7e-687">Value</span></span>|
|---|---|
|[<span data-ttu-id="d1b7e-688">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d1b7e-688">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d1b7e-689">1.5</span><span class="sxs-lookup"><span data-stu-id="d1b7e-689">1.5</span></span> |
|[<span data-ttu-id="d1b7e-690">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d1b7e-690">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d1b7e-691">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d1b7e-691">ReadItem</span></span> |
|[<span data-ttu-id="d1b7e-692">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d1b7e-692">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d1b7e-693">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d1b7e-693">Compose or Read</span></span>|
