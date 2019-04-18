---
title: Office のメールボックス-プレビュー要件セット
description: ''
ms.date: 04/17/2019
localization_priority: Normal
ms.openlocfilehash: 557dedf3943be12fbb9e384873d0b9079b251c2f
ms.sourcegitcommit: 6d375518c119d09c8d3fb5f0cc4583ba5b20ac03
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/18/2019
ms.locfileid: "31914334"
---
# <a name="mailbox"></a><span data-ttu-id="212a2-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="212a2-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="212a2-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="212a2-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="212a2-104">Microsoft Outlook と Microsoft Outlook on the web の Outlook アドイン オブジェクト モデルへのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="212a2-104">Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="212a2-105">要件</span><span class="sxs-lookup"><span data-stu-id="212a2-105">Requirements</span></span>

|<span data-ttu-id="212a2-106">要件</span><span class="sxs-lookup"><span data-stu-id="212a2-106">Requirement</span></span>| <span data-ttu-id="212a2-107">値</span><span class="sxs-lookup"><span data-stu-id="212a2-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="212a2-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="212a2-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="212a2-109">1.0</span><span class="sxs-lookup"><span data-stu-id="212a2-109">1.0</span></span>|
|[<span data-ttu-id="212a2-110">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="212a2-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="212a2-111">制限あり</span><span class="sxs-lookup"><span data-stu-id="212a2-111">Restricted</span></span>|
|[<span data-ttu-id="212a2-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="212a2-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="212a2-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="212a2-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="212a2-114">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="212a2-114">Members and methods</span></span>

| <span data-ttu-id="212a2-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="212a2-115">Member</span></span> | <span data-ttu-id="212a2-116">種類</span><span class="sxs-lookup"><span data-stu-id="212a2-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="212a2-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="212a2-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="212a2-118">メンバー</span><span class="sxs-lookup"><span data-stu-id="212a2-118">Member</span></span> |
| [<span data-ttu-id="212a2-119">masterCategories</span><span class="sxs-lookup"><span data-stu-id="212a2-119">masterCategories</span></span>](#mastercategories-mastercategories) | <span data-ttu-id="212a2-120">メンバー</span><span class="sxs-lookup"><span data-stu-id="212a2-120">Member</span></span> |
| [<span data-ttu-id="212a2-121">restUrl</span><span class="sxs-lookup"><span data-stu-id="212a2-121">restUrl</span></span>](#resturl-string) | <span data-ttu-id="212a2-122">メンバー</span><span class="sxs-lookup"><span data-stu-id="212a2-122">Member</span></span> |
| [<span data-ttu-id="212a2-123">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="212a2-123">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="212a2-124">メソッド</span><span class="sxs-lookup"><span data-stu-id="212a2-124">Method</span></span> |
| [<span data-ttu-id="212a2-125">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="212a2-125">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="212a2-126">メソッド</span><span class="sxs-lookup"><span data-stu-id="212a2-126">Method</span></span> |
| [<span data-ttu-id="212a2-127">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="212a2-127">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="212a2-128">メソッド</span><span class="sxs-lookup"><span data-stu-id="212a2-128">Method</span></span> |
| [<span data-ttu-id="212a2-129">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="212a2-129">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="212a2-130">メソッド</span><span class="sxs-lookup"><span data-stu-id="212a2-130">Method</span></span> |
| [<span data-ttu-id="212a2-131">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="212a2-131">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="212a2-132">メソッド</span><span class="sxs-lookup"><span data-stu-id="212a2-132">Method</span></span> |
| [<span data-ttu-id="212a2-133">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="212a2-133">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="212a2-134">メソッド</span><span class="sxs-lookup"><span data-stu-id="212a2-134">Method</span></span> |
| [<span data-ttu-id="212a2-135">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="212a2-135">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="212a2-136">メソッド</span><span class="sxs-lookup"><span data-stu-id="212a2-136">Method</span></span> |
| [<span data-ttu-id="212a2-137">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="212a2-137">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="212a2-138">メソッド</span><span class="sxs-lookup"><span data-stu-id="212a2-138">Method</span></span> |
| [<span data-ttu-id="212a2-139">displayNewMessageForm</span><span class="sxs-lookup"><span data-stu-id="212a2-139">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="212a2-140">メソッド</span><span class="sxs-lookup"><span data-stu-id="212a2-140">Method</span></span> |
| [<span data-ttu-id="212a2-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="212a2-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="212a2-142">メソッド</span><span class="sxs-lookup"><span data-stu-id="212a2-142">Method</span></span> |
| [<span data-ttu-id="212a2-143">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="212a2-143">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="212a2-144">メソッド</span><span class="sxs-lookup"><span data-stu-id="212a2-144">Method</span></span> |
| [<span data-ttu-id="212a2-145">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="212a2-145">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="212a2-146">メソッド</span><span class="sxs-lookup"><span data-stu-id="212a2-146">Method</span></span> |
| [<span data-ttu-id="212a2-147">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="212a2-147">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="212a2-148">メソッド</span><span class="sxs-lookup"><span data-stu-id="212a2-148">Method</span></span> |
| [<span data-ttu-id="212a2-149">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="212a2-149">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="212a2-150">メソッド</span><span class="sxs-lookup"><span data-stu-id="212a2-150">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="212a2-151">名前空間</span><span class="sxs-lookup"><span data-stu-id="212a2-151">Namespaces</span></span>

<span data-ttu-id="212a2-152">[diagnostics](Office.context.mailbox.diagnostics.md):Outlook アドインに診断情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="212a2-152">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="212a2-153">[item](Office.context.mailbox.item.md):Outlook アドインのメッセージや予定にアクセスするためのメソッドとプロパティを提供します。</span><span class="sxs-lookup"><span data-stu-id="212a2-153">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="212a2-154">[userProfile](Office.context.mailbox.userProfile.md):Outlook アドインのユーザーに関する情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="212a2-154">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="212a2-155">メンバー</span><span class="sxs-lookup"><span data-stu-id="212a2-155">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="212a2-156">ewsUrl :String</span><span class="sxs-lookup"><span data-stu-id="212a2-156">ewsUrl :String</span></span>

<span data-ttu-id="212a2-p101">このメール アカウントの Exchange Web サービス (EWS) エンドポイントの URL を取得します。読み取りモードのみです。</span><span class="sxs-lookup"><span data-stu-id="212a2-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="212a2-159">このメンバーは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="212a2-159">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="212a2-p102">
  `ewsUrl\` 値は、リモート サービスで、ユーザーのメールボックスに EWS 呼び出しを行うために使うことができます。たとえば、[選択したアイテムから添付ファイルを取得する](/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="212a2-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="212a2-162">アプリが閲覧モードで `ewsUrl` メンバーを呼び出すには、アプリのマニフェスト内に **ReadItem** アクセス許可が指定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="212a2-162">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="212a2-p103">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出してから、`ewsUrl` メンバーを使用する必要があります。アプリには、`saveAsync` メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="212a2-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="212a2-165">型</span><span class="sxs-lookup"><span data-stu-id="212a2-165">Type</span></span>

*   <span data-ttu-id="212a2-166">String</span><span class="sxs-lookup"><span data-stu-id="212a2-166">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="212a2-167">要件</span><span class="sxs-lookup"><span data-stu-id="212a2-167">Requirements</span></span>

|<span data-ttu-id="212a2-168">要件</span><span class="sxs-lookup"><span data-stu-id="212a2-168">Requirement</span></span>| <span data-ttu-id="212a2-169">値</span><span class="sxs-lookup"><span data-stu-id="212a2-169">Value</span></span>|
|---|---|
|[<span data-ttu-id="212a2-170">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="212a2-170">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="212a2-171">1.0</span><span class="sxs-lookup"><span data-stu-id="212a2-171">1.0</span></span>|
|[<span data-ttu-id="212a2-172">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="212a2-172">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="212a2-173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="212a2-173">ReadItem</span></span>|
|[<span data-ttu-id="212a2-174">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="212a2-174">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="212a2-175">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="212a2-175">Compose or Read</span></span>|

---
---

#### <a name="mastercategories-mastercategoriesjavascriptapioutlookofficemastercategories"></a><span data-ttu-id="212a2-176">mastercategories:[mastercategories](/javascript/api/outlook/office.mastercategories)</span><span class="sxs-lookup"><span data-stu-id="212a2-176">masterCategories :[MasterCategories](/javascript/api/outlook/office.mastercategories)</span></span>

<span data-ttu-id="212a2-177">このメールボックスのカテゴリマスターリストを管理するためのメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="212a2-177">Gets an object that provides methods to manage the categories master list on this mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="212a2-178">このメンバーは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="212a2-178">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="212a2-179">型</span><span class="sxs-lookup"><span data-stu-id="212a2-179">Type</span></span>

*   [<span data-ttu-id="212a2-180">MasterCategories</span><span class="sxs-lookup"><span data-stu-id="212a2-180">MasterCategories</span></span>](/javascript/api/outlook/office.mastercategories)

##### <a name="requirements"></a><span data-ttu-id="212a2-181">要件</span><span class="sxs-lookup"><span data-stu-id="212a2-181">Requirements</span></span>

|<span data-ttu-id="212a2-182">要件</span><span class="sxs-lookup"><span data-stu-id="212a2-182">Requirement</span></span>| <span data-ttu-id="212a2-183">値</span><span class="sxs-lookup"><span data-stu-id="212a2-183">Value</span></span>|
|---|---|
|[<span data-ttu-id="212a2-184">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="212a2-184">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="212a2-185">プレビュー</span><span class="sxs-lookup"><span data-stu-id="212a2-185">Preview</span></span> |
|[<span data-ttu-id="212a2-186">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="212a2-186">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="212a2-187">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="212a2-187">ReadWriteMailbox</span></span> |
|[<span data-ttu-id="212a2-188">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="212a2-188">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="212a2-189">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="212a2-189">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="212a2-190">例</span><span class="sxs-lookup"><span data-stu-id="212a2-190">Example</span></span>

<span data-ttu-id="212a2-191">この例では、このメールボックスのカテゴリマスターリストを取得します。</span><span class="sxs-lookup"><span data-stu-id="212a2-191">This example gets the categories master list for this mailbox.</span></span>

```javascript
Office.context.mailbox.masterCategories.getAsync(function (asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    console.log("Action failed with error: " + asyncResult.error.message);
  } else {
    console.log("Master categories: " + JSON.stringify(asyncResult.value));
  }
});
```

---
---

#### <a name="resturl-string"></a><span data-ttu-id="212a2-192">restUrl :String</span><span class="sxs-lookup"><span data-stu-id="212a2-192">restUrl :String</span></span>

<span data-ttu-id="212a2-193">この電子メール アカウントの REST エンドポイントの URL を取得します。</span><span class="sxs-lookup"><span data-stu-id="212a2-193">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="212a2-194">`restUrl` 値は、ユーザーのメールボックスに [REST API](/outlook/rest/) 呼び出しを行うために使用できます。</span><span class="sxs-lookup"><span data-stu-id="212a2-194">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="212a2-195">アプリが閲覧モードで `restUrl` メンバーを呼び出すには、アプリのマニフェスト内に **ReadItem** アクセス許可が指定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="212a2-195">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="212a2-p104">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出してから、`restUrl` メンバーを使用する必要があります。アプリには、`saveAsync` メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="212a2-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="212a2-198">型</span><span class="sxs-lookup"><span data-stu-id="212a2-198">Type</span></span>

*   <span data-ttu-id="212a2-199">String</span><span class="sxs-lookup"><span data-stu-id="212a2-199">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="212a2-200">要件</span><span class="sxs-lookup"><span data-stu-id="212a2-200">Requirements</span></span>

|<span data-ttu-id="212a2-201">要件</span><span class="sxs-lookup"><span data-stu-id="212a2-201">Requirement</span></span>| <span data-ttu-id="212a2-202">値</span><span class="sxs-lookup"><span data-stu-id="212a2-202">Value</span></span>|
|---|---|
|[<span data-ttu-id="212a2-203">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="212a2-203">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="212a2-204">1.5</span><span class="sxs-lookup"><span data-stu-id="212a2-204">1.5</span></span> |
|[<span data-ttu-id="212a2-205">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="212a2-205">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="212a2-206">ReadItem</span><span class="sxs-lookup"><span data-stu-id="212a2-206">ReadItem</span></span>|
|[<span data-ttu-id="212a2-207">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="212a2-207">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="212a2-208">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="212a2-208">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="212a2-209">メソッド</span><span class="sxs-lookup"><span data-stu-id="212a2-209">Methods</span></span>

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="212a2-210">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="212a2-210">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="212a2-211">サポートされているイベントのイベント ハンドラーを追加します。</span><span class="sxs-lookup"><span data-stu-id="212a2-211">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="212a2-212">現在、サポートされている`Office.EventType.ItemChanged`イベント`Office.EventType.OfficeThemeChanged`の種類はおよびです。</span><span class="sxs-lookup"><span data-stu-id="212a2-212">Currently, the supported event types are `Office.EventType.ItemChanged` and `Office.EventType.OfficeThemeChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="212a2-213">パラメーター</span><span class="sxs-lookup"><span data-stu-id="212a2-213">Parameters</span></span>

| <span data-ttu-id="212a2-214">名前</span><span class="sxs-lookup"><span data-stu-id="212a2-214">Name</span></span> | <span data-ttu-id="212a2-215">種類</span><span class="sxs-lookup"><span data-stu-id="212a2-215">Type</span></span> | <span data-ttu-id="212a2-216">属性</span><span class="sxs-lookup"><span data-stu-id="212a2-216">Attributes</span></span> | <span data-ttu-id="212a2-217">説明</span><span class="sxs-lookup"><span data-stu-id="212a2-217">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="212a2-218">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="212a2-218">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="212a2-219">ハンドラーを呼び出す必要のあるイベント。</span><span class="sxs-lookup"><span data-stu-id="212a2-219">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="212a2-220">関数</span><span class="sxs-lookup"><span data-stu-id="212a2-220">Function</span></span> || <span data-ttu-id="212a2-p105">イベントを処理する関数。関数は、オブジェクト リテラルである単一パラメーターを受け入れる必要があります。パラメーターの `type` プロパティは、`addHandlerAsync` に渡される `eventType` パラメーターと一致します。</span><span class="sxs-lookup"><span data-stu-id="212a2-p105">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="212a2-224">Object</span><span class="sxs-lookup"><span data-stu-id="212a2-224">Object</span></span> | <span data-ttu-id="212a2-225">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="212a2-225">&lt;optional&gt;</span></span> | <span data-ttu-id="212a2-226">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="212a2-226">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="212a2-227">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="212a2-227">Object</span></span> | <span data-ttu-id="212a2-228">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="212a2-228">&lt;optional&gt;</span></span> | <span data-ttu-id="212a2-229">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="212a2-229">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="212a2-230">function</span><span class="sxs-lookup"><span data-stu-id="212a2-230">function</span></span>| <span data-ttu-id="212a2-231">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="212a2-231">&lt;optional&gt;</span></span>|<span data-ttu-id="212a2-232">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="212a2-232">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="212a2-233">要件</span><span class="sxs-lookup"><span data-stu-id="212a2-233">Requirements</span></span>

|<span data-ttu-id="212a2-234">要件</span><span class="sxs-lookup"><span data-stu-id="212a2-234">Requirement</span></span>| <span data-ttu-id="212a2-235">値</span><span class="sxs-lookup"><span data-stu-id="212a2-235">Value</span></span>|
|---|---|
|[<span data-ttu-id="212a2-236">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="212a2-236">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="212a2-237">1.5</span><span class="sxs-lookup"><span data-stu-id="212a2-237">1.5</span></span> |
|[<span data-ttu-id="212a2-238">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="212a2-238">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="212a2-239">ReadItem</span><span class="sxs-lookup"><span data-stu-id="212a2-239">ReadItem</span></span> |
|[<span data-ttu-id="212a2-240">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="212a2-240">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="212a2-241">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="212a2-241">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="212a2-242">例</span><span class="sxs-lookup"><span data-stu-id="212a2-242">Example</span></span>

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

####  <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="212a2-243">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="212a2-243">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="212a2-244">REST 形式のアイテム ID を EWS 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="212a2-244">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="212a2-245">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="212a2-245">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="212a2-p106">REST API ([Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) や [Microsoft Graph](https://graph.microsoft.io/) など) で取得されたアイテム ID は、Exchange Web サービス (EWS) に使用される形式とは異なる形式を使用します。`convertToEwsId` メソッドは、REST 形式の ID を EWS 用の適切な形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="212a2-p106">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="212a2-248">パラメーター</span><span class="sxs-lookup"><span data-stu-id="212a2-248">Parameters</span></span>

|<span data-ttu-id="212a2-249">名前</span><span class="sxs-lookup"><span data-stu-id="212a2-249">Name</span></span>| <span data-ttu-id="212a2-250">種類</span><span class="sxs-lookup"><span data-stu-id="212a2-250">Type</span></span>| <span data-ttu-id="212a2-251">説明</span><span class="sxs-lookup"><span data-stu-id="212a2-251">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="212a2-252">String</span><span class="sxs-lookup"><span data-stu-id="212a2-252">String</span></span>|<span data-ttu-id="212a2-253">Outlook REST API 形式のアイテム ID</span><span class="sxs-lookup"><span data-stu-id="212a2-253">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="212a2-254">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="212a2-254">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion)|<span data-ttu-id="212a2-255">アイテム ID の取得に使用された Outlook REST API のバージョンを示す値。</span><span class="sxs-lookup"><span data-stu-id="212a2-255">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="212a2-256">要件</span><span class="sxs-lookup"><span data-stu-id="212a2-256">Requirements</span></span>

|<span data-ttu-id="212a2-257">要件</span><span class="sxs-lookup"><span data-stu-id="212a2-257">Requirement</span></span>| <span data-ttu-id="212a2-258">値</span><span class="sxs-lookup"><span data-stu-id="212a2-258">Value</span></span>|
|---|---|
|[<span data-ttu-id="212a2-259">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="212a2-259">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="212a2-260">1.3</span><span class="sxs-lookup"><span data-stu-id="212a2-260">1.3</span></span>|
|[<span data-ttu-id="212a2-261">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="212a2-261">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="212a2-262">制限あり</span><span class="sxs-lookup"><span data-stu-id="212a2-262">Restricted</span></span>|
|[<span data-ttu-id="212a2-263">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="212a2-263">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="212a2-264">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="212a2-264">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="212a2-265">戻り値:</span><span class="sxs-lookup"><span data-stu-id="212a2-265">Returns:</span></span>

<span data-ttu-id="212a2-266">型:String</span><span class="sxs-lookup"><span data-stu-id="212a2-266">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="212a2-267">例</span><span class="sxs-lookup"><span data-stu-id="212a2-267">Example</span></span>

```javascript
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

---
---

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime"></a><span data-ttu-id="212a2-268">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="212a2-268">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime)}</span></span>

<span data-ttu-id="212a2-269">クライアントのローカル時間で時間情報が含まれている辞書を取得します。</span><span class="sxs-lookup"><span data-stu-id="212a2-269">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="212a2-p107">Outlook 用メール アプリや Outlook Web App で使う日付と時刻では、異なるタイム ゾーンを使うことができます。Outlook では、クライアント コンピューターのタイム ゾーンを使います。Outlook Web App では、Exchange 管理センター (EAC) で設定されたタイム ゾーンを使います。ユーザー インターフェイスに表示される値が、常にユーザーが期待するタイム ゾーンと一致するように日付と時刻の値を処理する必要があります。</span><span class="sxs-lookup"><span data-stu-id="212a2-p107">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="212a2-p108">Outlook でメール アプリが実行されている場合、`convertToLocalClientTime` メソッドは、クライアント コンピューターのタイム ゾーンに設定された値のディクショナリ オブジェクトを返します。Outlook Web Apps でメール アプリが実行されている場合、`convertToLocalClientTime` メソッドは、EAC に指定されたタイム ゾーンに設定された値のディクショナリ オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="212a2-p108">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="212a2-275">パラメーター</span><span class="sxs-lookup"><span data-stu-id="212a2-275">Parameters</span></span>

|<span data-ttu-id="212a2-276">名前</span><span class="sxs-lookup"><span data-stu-id="212a2-276">Name</span></span>| <span data-ttu-id="212a2-277">種類</span><span class="sxs-lookup"><span data-stu-id="212a2-277">Type</span></span>| <span data-ttu-id="212a2-278">説明</span><span class="sxs-lookup"><span data-stu-id="212a2-278">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="212a2-279">Date</span><span class="sxs-lookup"><span data-stu-id="212a2-279">Date</span></span>|<span data-ttu-id="212a2-280">日付オブジェクト</span><span class="sxs-lookup"><span data-stu-id="212a2-280">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="212a2-281">要件</span><span class="sxs-lookup"><span data-stu-id="212a2-281">Requirements</span></span>

|<span data-ttu-id="212a2-282">要件</span><span class="sxs-lookup"><span data-stu-id="212a2-282">Requirement</span></span>| <span data-ttu-id="212a2-283">値</span><span class="sxs-lookup"><span data-stu-id="212a2-283">Value</span></span>|
|---|---|
|[<span data-ttu-id="212a2-284">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="212a2-284">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="212a2-285">1.0</span><span class="sxs-lookup"><span data-stu-id="212a2-285">1.0</span></span>|
|[<span data-ttu-id="212a2-286">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="212a2-286">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="212a2-287">ReadItem</span><span class="sxs-lookup"><span data-stu-id="212a2-287">ReadItem</span></span>|
|[<span data-ttu-id="212a2-288">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="212a2-288">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="212a2-289">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="212a2-289">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="212a2-290">戻り値:</span><span class="sxs-lookup"><span data-stu-id="212a2-290">Returns:</span></span>

<span data-ttu-id="212a2-291">型:[LocalClientTime](/javascript/api/outlook/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="212a2-291">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime)</span></span>

---
---

####  <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="212a2-292">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="212a2-292">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="212a2-293">EWS 形式のアイテム ID を REST 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="212a2-293">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="212a2-294">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="212a2-294">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="212a2-p109">EWS または `itemId` プロパティで取得されるアイテム ID は、REST API ([Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) や [Microsoft Graph](https://graph.microsoft.io/) など) に使用される形式とは異なる形式を使用します。`convertToRestId` メソッドは、EWS 形式の ID を REST 用の適切な形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="212a2-p109">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="212a2-297">パラメーター</span><span class="sxs-lookup"><span data-stu-id="212a2-297">Parameters</span></span>

|<span data-ttu-id="212a2-298">名前</span><span class="sxs-lookup"><span data-stu-id="212a2-298">Name</span></span>| <span data-ttu-id="212a2-299">種類</span><span class="sxs-lookup"><span data-stu-id="212a2-299">Type</span></span>| <span data-ttu-id="212a2-300">説明</span><span class="sxs-lookup"><span data-stu-id="212a2-300">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="212a2-301">String</span><span class="sxs-lookup"><span data-stu-id="212a2-301">String</span></span>|<span data-ttu-id="212a2-302">Exchange Web サービス (EWS) 形式のアイテム ID</span><span class="sxs-lookup"><span data-stu-id="212a2-302">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="212a2-303">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="212a2-303">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion)|<span data-ttu-id="212a2-304">変換後の ID を使用する Outlook REST API のバージョンを示す値。</span><span class="sxs-lookup"><span data-stu-id="212a2-304">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="212a2-305">要件</span><span class="sxs-lookup"><span data-stu-id="212a2-305">Requirements</span></span>

|<span data-ttu-id="212a2-306">要件</span><span class="sxs-lookup"><span data-stu-id="212a2-306">Requirement</span></span>| <span data-ttu-id="212a2-307">値</span><span class="sxs-lookup"><span data-stu-id="212a2-307">Value</span></span>|
|---|---|
|[<span data-ttu-id="212a2-308">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="212a2-308">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="212a2-309">1.3</span><span class="sxs-lookup"><span data-stu-id="212a2-309">1.3</span></span>|
|[<span data-ttu-id="212a2-310">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="212a2-310">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="212a2-311">制限あり</span><span class="sxs-lookup"><span data-stu-id="212a2-311">Restricted</span></span>|
|[<span data-ttu-id="212a2-312">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="212a2-312">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="212a2-313">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="212a2-313">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="212a2-314">戻り値:</span><span class="sxs-lookup"><span data-stu-id="212a2-314">Returns:</span></span>

<span data-ttu-id="212a2-315">型:String</span><span class="sxs-lookup"><span data-stu-id="212a2-315">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="212a2-316">例</span><span class="sxs-lookup"><span data-stu-id="212a2-316">Example</span></span>

```javascript
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

---
---

####  <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="212a2-317">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="212a2-317">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="212a2-318">時間情報が含まれているディクショナリから日付オブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="212a2-318">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="212a2-319">`convertToUtcClientTime` メソッドは、ローカルの日付と時刻を含むディクショナリを、ローカルの日付と時刻の正しい値を持つ日付オブジェクトに変換します。</span><span class="sxs-lookup"><span data-stu-id="212a2-319">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="212a2-320">パラメーター</span><span class="sxs-lookup"><span data-stu-id="212a2-320">Parameters</span></span>

|<span data-ttu-id="212a2-321">名前</span><span class="sxs-lookup"><span data-stu-id="212a2-321">Name</span></span>| <span data-ttu-id="212a2-322">種類</span><span class="sxs-lookup"><span data-stu-id="212a2-322">Type</span></span>| <span data-ttu-id="212a2-323">説明</span><span class="sxs-lookup"><span data-stu-id="212a2-323">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="212a2-324">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="212a2-324">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime)|<span data-ttu-id="212a2-325">変換するローカル時刻の値。</span><span class="sxs-lookup"><span data-stu-id="212a2-325">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="212a2-326">要件</span><span class="sxs-lookup"><span data-stu-id="212a2-326">Requirements</span></span>

|<span data-ttu-id="212a2-327">要件</span><span class="sxs-lookup"><span data-stu-id="212a2-327">Requirement</span></span>| <span data-ttu-id="212a2-328">値</span><span class="sxs-lookup"><span data-stu-id="212a2-328">Value</span></span>|
|---|---|
|[<span data-ttu-id="212a2-329">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="212a2-329">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="212a2-330">1.0</span><span class="sxs-lookup"><span data-stu-id="212a2-330">1.0</span></span>|
|[<span data-ttu-id="212a2-331">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="212a2-331">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="212a2-332">ReadItem</span><span class="sxs-lookup"><span data-stu-id="212a2-332">ReadItem</span></span>|
|[<span data-ttu-id="212a2-333">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="212a2-333">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="212a2-334">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="212a2-334">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="212a2-335">戻り値:</span><span class="sxs-lookup"><span data-stu-id="212a2-335">Returns:</span></span>

<span data-ttu-id="212a2-336">時間が UTC で表現された日付オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="212a2-336">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="212a2-337">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="212a2-337">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="212a2-338">日付</span><span class="sxs-lookup"><span data-stu-id="212a2-338">Date</span></span></dd>

</dl>

---
---

####  <a name="displayappointmentformitemid"></a><span data-ttu-id="212a2-339">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="212a2-339">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="212a2-340">既存の予定を表示します。</span><span class="sxs-lookup"><span data-stu-id="212a2-340">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="212a2-341">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="212a2-341">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="212a2-342">`displayAppointmentForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存の予定を開きます。</span><span class="sxs-lookup"><span data-stu-id="212a2-342">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="212a2-p110">Outlook for Mac では、この方法を使って、定期的な系列の一部ではない単一の予定や定期的な系列のマスター予定を表示できます。ただし、系列のインスタンスは表示できません。これは、Outlook for Mac においては定期的な系列のインスタンスのプロパティ (アイテム ID を含む) にアクセスできないためです。</span><span class="sxs-lookup"><span data-stu-id="212a2-p110">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="212a2-345">Outlook Web App では、このメソッドは指定されたフォームの本文が 32 KB 以下の文字数の場合にフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="212a2-345">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="212a2-346">指定のアイテム識別子が既存の予定を表していない場合は、クライアント コンピューターまたはデバイスで空のウィンドウが開き、エラー メッセージは返されません。</span><span class="sxs-lookup"><span data-stu-id="212a2-346">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="212a2-347">パラメーター</span><span class="sxs-lookup"><span data-stu-id="212a2-347">Parameters</span></span>

|<span data-ttu-id="212a2-348">名前</span><span class="sxs-lookup"><span data-stu-id="212a2-348">Name</span></span>| <span data-ttu-id="212a2-349">種類</span><span class="sxs-lookup"><span data-stu-id="212a2-349">Type</span></span>| <span data-ttu-id="212a2-350">説明</span><span class="sxs-lookup"><span data-stu-id="212a2-350">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="212a2-351">String</span><span class="sxs-lookup"><span data-stu-id="212a2-351">String</span></span>|<span data-ttu-id="212a2-352">既存の予定の Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="212a2-352">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="212a2-353">要件</span><span class="sxs-lookup"><span data-stu-id="212a2-353">Requirements</span></span>

|<span data-ttu-id="212a2-354">要件</span><span class="sxs-lookup"><span data-stu-id="212a2-354">Requirement</span></span>| <span data-ttu-id="212a2-355">値</span><span class="sxs-lookup"><span data-stu-id="212a2-355">Value</span></span>|
|---|---|
|[<span data-ttu-id="212a2-356">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="212a2-356">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="212a2-357">1.0</span><span class="sxs-lookup"><span data-stu-id="212a2-357">1.0</span></span>|
|[<span data-ttu-id="212a2-358">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="212a2-358">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="212a2-359">ReadItem</span><span class="sxs-lookup"><span data-stu-id="212a2-359">ReadItem</span></span>|
|[<span data-ttu-id="212a2-360">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="212a2-360">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="212a2-361">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="212a2-361">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="212a2-362">例</span><span class="sxs-lookup"><span data-stu-id="212a2-362">Example</span></span>

```javascript
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

---
---

####  <a name="displaymessageformitemid"></a><span data-ttu-id="212a2-363">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="212a2-363">displayMessageForm(itemId)</span></span>

<span data-ttu-id="212a2-364">既存のメッセージを表示します。</span><span class="sxs-lookup"><span data-stu-id="212a2-364">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="212a2-365">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="212a2-365">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="212a2-366">`displayMessageForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存のメッセージを開きます。</span><span class="sxs-lookup"><span data-stu-id="212a2-366">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="212a2-367">Outlook Web App では、このメソッドは指定されたフォームの本文が 32 KB 以下の文字数の場合にフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="212a2-367">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="212a2-368">指定のアイテム識別子が既存のメッセージを表していない場合は、クライアント コンピューターにはメッセージは表示されず、エラー メッセージも返されません。</span><span class="sxs-lookup"><span data-stu-id="212a2-368">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="212a2-p111">予定を表す `itemId` を含む `displayMessageForm` を使用しないでください。既存の予定を表示するには、`displayAppointmentForm` メソッドを使用します。新しい予定を作成するフォームを表示するには、`displayNewAppointmentForm` メソッドを使用します。</span><span class="sxs-lookup"><span data-stu-id="212a2-p111">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="212a2-371">パラメーター</span><span class="sxs-lookup"><span data-stu-id="212a2-371">Parameters</span></span>

|<span data-ttu-id="212a2-372">名前</span><span class="sxs-lookup"><span data-stu-id="212a2-372">Name</span></span>| <span data-ttu-id="212a2-373">型</span><span class="sxs-lookup"><span data-stu-id="212a2-373">Type</span></span>| <span data-ttu-id="212a2-374">説明</span><span class="sxs-lookup"><span data-stu-id="212a2-374">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="212a2-375">String</span><span class="sxs-lookup"><span data-stu-id="212a2-375">String</span></span>|<span data-ttu-id="212a2-376">既存のメッセージの Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="212a2-376">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="212a2-377">要件</span><span class="sxs-lookup"><span data-stu-id="212a2-377">Requirements</span></span>

|<span data-ttu-id="212a2-378">要件</span><span class="sxs-lookup"><span data-stu-id="212a2-378">Requirement</span></span>| <span data-ttu-id="212a2-379">値</span><span class="sxs-lookup"><span data-stu-id="212a2-379">Value</span></span>|
|---|---|
|[<span data-ttu-id="212a2-380">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="212a2-380">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="212a2-381">1.0</span><span class="sxs-lookup"><span data-stu-id="212a2-381">1.0</span></span>|
|[<span data-ttu-id="212a2-382">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="212a2-382">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="212a2-383">ReadItem</span><span class="sxs-lookup"><span data-stu-id="212a2-383">ReadItem</span></span>|
|[<span data-ttu-id="212a2-384">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="212a2-384">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="212a2-385">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="212a2-385">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="212a2-386">例</span><span class="sxs-lookup"><span data-stu-id="212a2-386">Example</span></span>

```javascript
Office.context.mailbox.displayMessageForm(messageId);
```

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="212a2-387">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="212a2-387">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="212a2-388">新しい予定を作成するためのフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="212a2-388">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="212a2-389">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="212a2-389">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="212a2-p112">`displayNewAppointmentForm` メソッドを使用すると、ユーザーが新しい予定または会議を作成できるフォームが開きます。パラメーターを指定すると、予定のフォーム フィールドにパラメーターの内容が自動的に設定されます。</span><span class="sxs-lookup"><span data-stu-id="212a2-p112">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="212a2-p113">このメソッドは、Outlook Web App と OWA for Devices において、出席者フィールドが含まれるフォームを必ず表示します。入力引数として出席者を指定しないと、このメソッドにより **[保存]** ボタンのあるフォームが表示されます。出席者を指定した場合には、フォームにその出席者と **[送信]** ボタンが表示されます。</span><span class="sxs-lookup"><span data-stu-id="212a2-p113">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="212a2-p114">Outlook リッチ クライアントと Outlook RT で、`requiredAttendees`、`optionalAttendees`、または `resources` パラメーターに出席者またはリソースを指定し、このメソッドを実行すると、**[送信]** ボタンがある会議フォームが表示されます。受信者を指定せずにこのメソッドを実行すると、**[保存して閉じる]** ボタンがある予定フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="212a2-p114">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="212a2-397">パラメーターのいずれかが指定のサイズ制限を超える場合、または不明なパラメーター名が指定されている場合は、例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="212a2-397">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="212a2-398">パラメーター</span><span class="sxs-lookup"><span data-stu-id="212a2-398">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="212a2-399">すべてのパラメーターは省略可能です。</span><span class="sxs-lookup"><span data-stu-id="212a2-399">All parameters are optional.</span></span>

|<span data-ttu-id="212a2-400">名前</span><span class="sxs-lookup"><span data-stu-id="212a2-400">Name</span></span>| <span data-ttu-id="212a2-401">型</span><span class="sxs-lookup"><span data-stu-id="212a2-401">Type</span></span>| <span data-ttu-id="212a2-402">説明</span><span class="sxs-lookup"><span data-stu-id="212a2-402">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="212a2-403">Object</span><span class="sxs-lookup"><span data-stu-id="212a2-403">Object</span></span> | <span data-ttu-id="212a2-404">新しい予定を記述するパラメーターのディクショナリ。</span><span class="sxs-lookup"><span data-stu-id="212a2-404">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="212a2-405">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="212a2-405">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="212a2-p115">予定に必要な各出席者について、メール アドレスを含む文字列の配列、または `EmailAddressDetails` オブジェクトを含む配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="212a2-p115">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="212a2-408">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="212a2-408">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="212a2-p116">予定の各任意出席者について、メール アドレスを含む文字列の配列、または `EmailAddressDetails` オブジェクトを含む配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="212a2-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="212a2-411">日付</span><span class="sxs-lookup"><span data-stu-id="212a2-411">Date</span></span> | <span data-ttu-id="212a2-412">予定の開始日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="212a2-412">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="212a2-413">日付</span><span class="sxs-lookup"><span data-stu-id="212a2-413">Date</span></span> | <span data-ttu-id="212a2-414">予定の終了日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="212a2-414">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="212a2-415">String</span><span class="sxs-lookup"><span data-stu-id="212a2-415">String</span></span> | <span data-ttu-id="212a2-p117">予定の場所を含む文字列。文字列は最大 255 文字に制限されます。</span><span class="sxs-lookup"><span data-stu-id="212a2-p117">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="212a2-418">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="212a2-418">Array.&lt;String&gt;</span></span> | <span data-ttu-id="212a2-p118">予定に必要なリソースを含む文字列の配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="212a2-p118">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="212a2-421">String</span><span class="sxs-lookup"><span data-stu-id="212a2-421">String</span></span> | <span data-ttu-id="212a2-p119">予定の件名を含む文字列です。文字列は最大 255 文字に制限されます。</span><span class="sxs-lookup"><span data-stu-id="212a2-p119">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="212a2-424">String</span><span class="sxs-lookup"><span data-stu-id="212a2-424">String</span></span> | <span data-ttu-id="212a2-p120">予定の本文。本文の内容は、最大サイズが 32 KB に制限されます。</span><span class="sxs-lookup"><span data-stu-id="212a2-p120">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="212a2-427">要件</span><span class="sxs-lookup"><span data-stu-id="212a2-427">Requirements</span></span>

|<span data-ttu-id="212a2-428">要件</span><span class="sxs-lookup"><span data-stu-id="212a2-428">Requirement</span></span>| <span data-ttu-id="212a2-429">値</span><span class="sxs-lookup"><span data-stu-id="212a2-429">Value</span></span>|
|---|---|
|[<span data-ttu-id="212a2-430">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="212a2-430">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="212a2-431">1.0</span><span class="sxs-lookup"><span data-stu-id="212a2-431">1.0</span></span>|
|[<span data-ttu-id="212a2-432">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="212a2-432">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="212a2-433">ReadItem</span><span class="sxs-lookup"><span data-stu-id="212a2-433">ReadItem</span></span>|
|[<span data-ttu-id="212a2-434">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="212a2-434">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="212a2-435">読み取り</span><span class="sxs-lookup"><span data-stu-id="212a2-435">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="212a2-436">例</span><span class="sxs-lookup"><span data-stu-id="212a2-436">Example</span></span>

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

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="212a2-437">displayNewMessageForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="212a2-437">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="212a2-438">新しいメッセージを作成するためのフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="212a2-438">Displays a form for creating a new message.</span></span>

<span data-ttu-id="212a2-p121">The `displayNewMessageForm` method opens a form that enables the user to create a new message. If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span><span class="sxs-lookup"><span data-stu-id="212a2-p121">The `displayNewMessageForm` method opens a form that enables the user to create a new message. If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="212a2-441">パラメーターのいずれかが指定のサイズ制限を超える場合、または不明なパラメーター名が指定されている場合は、例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="212a2-441">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="212a2-442">パラメーター</span><span class="sxs-lookup"><span data-stu-id="212a2-442">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="212a2-443">すべてのパラメーターは省略可能です。</span><span class="sxs-lookup"><span data-stu-id="212a2-443">All parameters are optional.</span></span>

|<span data-ttu-id="212a2-444">名前</span><span class="sxs-lookup"><span data-stu-id="212a2-444">Name</span></span>| <span data-ttu-id="212a2-445">型</span><span class="sxs-lookup"><span data-stu-id="212a2-445">Type</span></span>| <span data-ttu-id="212a2-446">説明</span><span class="sxs-lookup"><span data-stu-id="212a2-446">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="212a2-447">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="212a2-447">Object</span></span> | <span data-ttu-id="212a2-448">新しいメッセージを記述するパラメータの辞書。</span><span class="sxs-lookup"><span data-stu-id="212a2-448">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="212a2-449">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="212a2-449">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="212a2-p122">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line. The array is limited to a maximum of 100 entries.</span><span class="sxs-lookup"><span data-stu-id="212a2-p122">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="212a2-452">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="212a2-452">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="212a2-p123">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line. The array is limited to a maximum of 100 entries.</span><span class="sxs-lookup"><span data-stu-id="212a2-p123">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="212a2-455">配列。&lt;文字列&gt; | 配列。&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="212a2-455">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="212a2-p124">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line. The array is limited to a maximum of 100 entries.</span><span class="sxs-lookup"><span data-stu-id="212a2-p124">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="212a2-458">String</span><span class="sxs-lookup"><span data-stu-id="212a2-458">String</span></span> | <span data-ttu-id="212a2-p125">A string containing the subject of the message. The string is limited to a maximum of 255 characters.</span><span class="sxs-lookup"><span data-stu-id="212a2-p125">A string containing the subject of the message. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="212a2-461">String</span><span class="sxs-lookup"><span data-stu-id="212a2-461">String</span></span> | <span data-ttu-id="212a2-p126">The HTML body of the message. The body content is limited to a maximum size of 32 KB.</span><span class="sxs-lookup"><span data-stu-id="212a2-p126">The HTML body of the message. The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="212a2-464">配列。&lt;オブジェクト&gt;</span><span class="sxs-lookup"><span data-stu-id="212a2-464">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="212a2-465">添付ファイルまたは添付アイテムである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="212a2-465">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="212a2-466">String</span><span class="sxs-lookup"><span data-stu-id="212a2-466">String</span></span> | <span data-ttu-id="212a2-p127">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="212a2-p127">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="212a2-469">String</span><span class="sxs-lookup"><span data-stu-id="212a2-469">String</span></span> | <span data-ttu-id="212a2-470">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="212a2-470">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="212a2-471">文字列</span><span class="sxs-lookup"><span data-stu-id="212a2-471">String</span></span> | <span data-ttu-id="212a2-p128">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="212a2-p128">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="212a2-474">ブール値</span><span class="sxs-lookup"><span data-stu-id="212a2-474">Boolean</span></span> | <span data-ttu-id="212a2-p129">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="212a2-p129">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="212a2-477">String</span><span class="sxs-lookup"><span data-stu-id="212a2-477">String</span></span> | <span data-ttu-id="212a2-p130">Only used if `type` is set to `item`. The EWS item id of the existing e-mail you want to attach to the new message. This is a string up to 100 characters.</span><span class="sxs-lookup"><span data-stu-id="212a2-p130">Only used if `type` is set to `item`. The EWS item id of the existing e-mail you want to attach to the new message. This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="212a2-481">要件</span><span class="sxs-lookup"><span data-stu-id="212a2-481">Requirements</span></span>

|<span data-ttu-id="212a2-482">要件</span><span class="sxs-lookup"><span data-stu-id="212a2-482">Requirement</span></span>| <span data-ttu-id="212a2-483">値</span><span class="sxs-lookup"><span data-stu-id="212a2-483">Value</span></span>|
|---|---|
|[<span data-ttu-id="212a2-484">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="212a2-484">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="212a2-485">1.6</span><span class="sxs-lookup"><span data-stu-id="212a2-485">1.6</span></span> |
|[<span data-ttu-id="212a2-486">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="212a2-486">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="212a2-487">ReadItem</span><span class="sxs-lookup"><span data-stu-id="212a2-487">ReadItem</span></span>|
|[<span data-ttu-id="212a2-488">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="212a2-488">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="212a2-489">読み取り</span><span class="sxs-lookup"><span data-stu-id="212a2-489">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="212a2-490">例</span><span class="sxs-lookup"><span data-stu-id="212a2-490">Example</span></span>

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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="212a2-491">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="212a2-491">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="212a2-492">REST API または Exchange Web サービスを呼び出すために使用するトークンを含む文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="212a2-492">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="212a2-p131">`getCallbackTokenAsync` メソッドは、ユーザーのメールボックスをホストする Exchange Server から不透明なトークンを取得する非同期の呼び出しを行います。コールバック トークンの有効期間は 5 分です。</span><span class="sxs-lookup"><span data-stu-id="212a2-p131">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="212a2-495">可能であれば、アドインでは Exchange Web サービスの代わりに REST API を使用することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="212a2-495">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span>

<span data-ttu-id="212a2-496">**REST トークン**</span><span class="sxs-lookup"><span data-stu-id="212a2-496">**REST Tokens**</span></span>

<span data-ttu-id="212a2-p132">REST トークンが要求された場合 (`options.isRest = true`)、結果トークンは Exchange Web サービスの呼び出しを認証するためには機能しません。アドインがマニフェストで [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) アクセス許可を指定していない限り、トークンの範囲は現在のアイテムとその添付ファイルへの読み取り専用アクセスに制限されます。`ReadWriteMailbox` アクセス許可が指定されている場合は、結果トークンは、メールを送信する機能など、メール、カレンダー、連絡先への読み取り/書き込みアクセスを付与します。</span><span class="sxs-lookup"><span data-stu-id="212a2-p132">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="212a2-500">アドインでは、`restUrl` プロパティを使用して、REST API 呼び出しを行うときに使用する正しい URL を決定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="212a2-500">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="212a2-501">**EWS トークン**</span><span class="sxs-lookup"><span data-stu-id="212a2-501">**EWS Tokens**</span></span>

<span data-ttu-id="212a2-p133">EWS トークンが要求された場合 (`options.isRest = false`)、結果トークンは REST API 呼び出しを認証するためには機能しません。トークンの範囲は、現在のアイテムへのアクセスに制限されます。</span><span class="sxs-lookup"><span data-stu-id="212a2-p133">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="212a2-504">アドインでは、`ewsUrl` プロパティを使用して、EWS 呼び出しを行うときに使用する正しい URL を決定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="212a2-504">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="212a2-505">パラメーター</span><span class="sxs-lookup"><span data-stu-id="212a2-505">Parameters</span></span>

|<span data-ttu-id="212a2-506">名前</span><span class="sxs-lookup"><span data-stu-id="212a2-506">Name</span></span>| <span data-ttu-id="212a2-507">型</span><span class="sxs-lookup"><span data-stu-id="212a2-507">Type</span></span>| <span data-ttu-id="212a2-508">属性</span><span class="sxs-lookup"><span data-stu-id="212a2-508">Attributes</span></span>| <span data-ttu-id="212a2-509">説明</span><span class="sxs-lookup"><span data-stu-id="212a2-509">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="212a2-510">Object</span><span class="sxs-lookup"><span data-stu-id="212a2-510">Object</span></span> | <span data-ttu-id="212a2-511">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="212a2-511">&lt;optional&gt;</span></span> | <span data-ttu-id="212a2-512">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="212a2-512">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="212a2-513">Boolean</span><span class="sxs-lookup"><span data-stu-id="212a2-513">Boolean</span></span> |  <span data-ttu-id="212a2-514">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="212a2-514">&lt;optional&gt;</span></span> | <span data-ttu-id="212a2-p134">提供されたトークンを Outlook REST API または Exchange Web サービスに使用するかどうかを決定します。既定値は、`false` です。</span><span class="sxs-lookup"><span data-stu-id="212a2-p134">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="212a2-517">Object</span><span class="sxs-lookup"><span data-stu-id="212a2-517">Object</span></span> |  <span data-ttu-id="212a2-518">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="212a2-518">&lt;optional&gt;</span></span> | <span data-ttu-id="212a2-519">非同期メソッドに渡される状態データ。</span><span class="sxs-lookup"><span data-stu-id="212a2-519">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="212a2-520">function</span><span class="sxs-lookup"><span data-stu-id="212a2-520">function</span></span>||<span data-ttu-id="212a2-p135">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="212a2-p135">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="212a2-523">要件</span><span class="sxs-lookup"><span data-stu-id="212a2-523">Requirements</span></span>

|<span data-ttu-id="212a2-524">要件</span><span class="sxs-lookup"><span data-stu-id="212a2-524">Requirement</span></span>| <span data-ttu-id="212a2-525">値</span><span class="sxs-lookup"><span data-stu-id="212a2-525">Value</span></span>|
|---|---|
|[<span data-ttu-id="212a2-526">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="212a2-526">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="212a2-527">1.5</span><span class="sxs-lookup"><span data-stu-id="212a2-527">1.5</span></span> |
|[<span data-ttu-id="212a2-528">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="212a2-528">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="212a2-529">ReadItem</span><span class="sxs-lookup"><span data-stu-id="212a2-529">ReadItem</span></span>|
|[<span data-ttu-id="212a2-530">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="212a2-530">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="212a2-531">新規作成と閲覧</span><span class="sxs-lookup"><span data-stu-id="212a2-531">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="212a2-532">例</span><span class="sxs-lookup"><span data-stu-id="212a2-532">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="212a2-533">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="212a2-533">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="212a2-534">Exchange Server から添付ファイルやアイテムを取得するために使うトークンを含む文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="212a2-534">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="212a2-p136">`getCallbackTokenAsync` メソッドは、ユーザーのメールボックスをホストする Exchange Server から不透明なトークンを取得する非同期の呼び出しを行います。コールバック トークンの有効期間は 5 分です。</span><span class="sxs-lookup"><span data-stu-id="212a2-p136">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="212a2-p137">トークンと予定の識別子またはアイテムの識別子をサードパーティ システムに渡すことができます。サードパーティ システムは、トークンをベアラー承認トークンとして使用し、Exchange Web サービス (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) または [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) 操作を呼び出して、添付ファイルまたはアイテムを返します。たとえば、[選択したアイテムから添付ファイルを取得する](/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="212a2-p137">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="212a2-540">アプリが閲覧モードで `getCallbackTokenAsync` メソッドを呼び出すには、アプリのマニフェスト内に **ReadItem** アクセス許可が指定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="212a2-540">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="212a2-p138">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出してアイテムの識別子を `getCallbackTokenAsync` メソッドに渡す必要があります。アプリには、`saveAsync` メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="212a2-p138">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="212a2-543">パラメーター</span><span class="sxs-lookup"><span data-stu-id="212a2-543">Parameters</span></span>

|<span data-ttu-id="212a2-544">名前</span><span class="sxs-lookup"><span data-stu-id="212a2-544">Name</span></span>| <span data-ttu-id="212a2-545">型</span><span class="sxs-lookup"><span data-stu-id="212a2-545">Type</span></span>| <span data-ttu-id="212a2-546">属性</span><span class="sxs-lookup"><span data-stu-id="212a2-546">Attributes</span></span>| <span data-ttu-id="212a2-547">説明</span><span class="sxs-lookup"><span data-stu-id="212a2-547">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="212a2-548">function</span><span class="sxs-lookup"><span data-stu-id="212a2-548">function</span></span>||<span data-ttu-id="212a2-p139">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="212a2-p139">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="212a2-551">Object</span><span class="sxs-lookup"><span data-stu-id="212a2-551">Object</span></span>| <span data-ttu-id="212a2-552">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="212a2-552">&lt;optional&gt;</span></span>|<span data-ttu-id="212a2-553">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="212a2-553">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="212a2-554">要件</span><span class="sxs-lookup"><span data-stu-id="212a2-554">Requirements</span></span>

|<span data-ttu-id="212a2-555">要件</span><span class="sxs-lookup"><span data-stu-id="212a2-555">Requirement</span></span>| <span data-ttu-id="212a2-556">値</span><span class="sxs-lookup"><span data-stu-id="212a2-556">Value</span></span>|
|---|---|
|[<span data-ttu-id="212a2-557">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="212a2-557">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="212a2-558">1.3</span><span class="sxs-lookup"><span data-stu-id="212a2-558">1.3</span></span>|
|[<span data-ttu-id="212a2-559">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="212a2-559">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="212a2-560">ReadItem</span><span class="sxs-lookup"><span data-stu-id="212a2-560">ReadItem</span></span>|
|[<span data-ttu-id="212a2-561">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="212a2-561">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="212a2-562">新規作成と閲覧</span><span class="sxs-lookup"><span data-stu-id="212a2-562">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="212a2-563">例</span><span class="sxs-lookup"><span data-stu-id="212a2-563">Example</span></span>

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

####  <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="212a2-564">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="212a2-564">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="212a2-565">ユーザーと Office アドインを識別するトークンを取得します。</span><span class="sxs-lookup"><span data-stu-id="212a2-565">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="212a2-566">`getUserIdentityTokenAsync` メソッドは、[アドインとユーザーをサード パーティのシステムで識別して認証](/outlook/add-ins/authentication)することのできるトークンを返します。</span><span class="sxs-lookup"><span data-stu-id="212a2-566">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="212a2-567">パラメーター</span><span class="sxs-lookup"><span data-stu-id="212a2-567">Parameters</span></span>

|<span data-ttu-id="212a2-568">名前</span><span class="sxs-lookup"><span data-stu-id="212a2-568">Name</span></span>| <span data-ttu-id="212a2-569">型</span><span class="sxs-lookup"><span data-stu-id="212a2-569">Type</span></span>| <span data-ttu-id="212a2-570">属性</span><span class="sxs-lookup"><span data-stu-id="212a2-570">Attributes</span></span>| <span data-ttu-id="212a2-571">説明</span><span class="sxs-lookup"><span data-stu-id="212a2-571">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="212a2-572">関数</span><span class="sxs-lookup"><span data-stu-id="212a2-572">function</span></span>||<span data-ttu-id="212a2-573">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="212a2-573">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="212a2-574">トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="212a2-574">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="212a2-575">Object</span><span class="sxs-lookup"><span data-stu-id="212a2-575">Object</span></span>| <span data-ttu-id="212a2-576">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="212a2-576">&lt;optional&gt;</span></span>|<span data-ttu-id="212a2-577">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="212a2-577">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="212a2-578">要件</span><span class="sxs-lookup"><span data-stu-id="212a2-578">Requirements</span></span>

|<span data-ttu-id="212a2-579">要件</span><span class="sxs-lookup"><span data-stu-id="212a2-579">Requirement</span></span>| <span data-ttu-id="212a2-580">値</span><span class="sxs-lookup"><span data-stu-id="212a2-580">Value</span></span>|
|---|---|
|[<span data-ttu-id="212a2-581">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="212a2-581">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="212a2-582">1.0</span><span class="sxs-lookup"><span data-stu-id="212a2-582">1.0</span></span>|
|[<span data-ttu-id="212a2-583">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="212a2-583">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="212a2-584">ReadItem</span><span class="sxs-lookup"><span data-stu-id="212a2-584">ReadItem</span></span>|
|[<span data-ttu-id="212a2-585">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="212a2-585">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="212a2-586">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="212a2-586">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="212a2-587">例</span><span class="sxs-lookup"><span data-stu-id="212a2-587">Example</span></span>

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

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="212a2-588">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="212a2-588">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="212a2-589">ユーザーのメールボックスをホストしている Exchange サーバー上の Exchange Web サービス (EWS) のサービスに対して非同期の要求を行います。</span><span class="sxs-lookup"><span data-stu-id="212a2-589">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="212a2-590">このメソッドは、次のシナリオではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="212a2-590">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="212a2-591">Outlook for iOS または Outlook for Android を使用している場合</span><span class="sxs-lookup"><span data-stu-id="212a2-591">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="212a2-592">アドインが Gmail のメールボックスに読み込まれる場合</span><span class="sxs-lookup"><span data-stu-id="212a2-592">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="212a2-593">このような場合は、アドインでは [REST API を使用](/outlook/add-ins/use-rest-api)して、代わりにユーザーのメールボックスにアクセスする必要があります。</span><span class="sxs-lookup"><span data-stu-id="212a2-593">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="212a2-p140">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange. See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span><span class="sxs-lookup"><span data-stu-id="212a2-p140">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange. See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="212a2-596">`makeEwsRequestAsync` メソッドでは、フォルダー関連アイテムを要求できません。</span><span class="sxs-lookup"><span data-stu-id="212a2-596">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="212a2-597">XML 要求では UTF-8 エンコードを指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="212a2-597">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="212a2-p141">`makeEwsRequestAsync` メソッドを使用するには、アドインに **ReadWriteMailbox** アクセス許可が必要です。**ReadWriteMailbox** アクセス許可と、`makeEwsRequestAsync` メソッドで呼び出せる EWS 操作の使い方については、「[ユーザーのメールボックスへのメール アドイン アクセスのアクセス許可を指定する](/outlook/add-ins/understanding-outlook-add-in-permissions)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="212a2-p141">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="212a2-600">サーバー管理者は、クライアント アクセス サーバーの EWS ディレクトリで `OAuthAuthentication` を true に設定して、`makeEwsRequestAsync` メソッドで EWS 要求を行うことができるようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="212a2-600">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="212a2-601">バージョンの相違点</span><span class="sxs-lookup"><span data-stu-id="212a2-601">Version differences</span></span>

<span data-ttu-id="212a2-602">バージョン 15.0.4535.1004 より前のバージョンの Outlook で実行しているメール アプリで `makeEwsRequestAsync` メソッドを使う場合は、エンコード値を `ISO-8859-1` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="212a2-602">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="212a2-p142">Outlook on the web でメール アプリを実行している場合は、エンコード値を設定する必要はありません。mailbox.diagnostics.hostName プロパティを使って、メール アプリを Outlook で実行しているのか、Outlook on the web で実行しているのかを確認できます。mailbox.diagnostics.hostVersion プロパティを使って、どのバージョンの Outlook を使って実行しているかを確認できます。</span><span class="sxs-lookup"><span data-stu-id="212a2-p142">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="212a2-606">パラメーター</span><span class="sxs-lookup"><span data-stu-id="212a2-606">Parameters</span></span>

|<span data-ttu-id="212a2-607">名前</span><span class="sxs-lookup"><span data-stu-id="212a2-607">Name</span></span>| <span data-ttu-id="212a2-608">型</span><span class="sxs-lookup"><span data-stu-id="212a2-608">Type</span></span>| <span data-ttu-id="212a2-609">属性</span><span class="sxs-lookup"><span data-stu-id="212a2-609">Attributes</span></span>| <span data-ttu-id="212a2-610">説明</span><span class="sxs-lookup"><span data-stu-id="212a2-610">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="212a2-611">String</span><span class="sxs-lookup"><span data-stu-id="212a2-611">String</span></span>||<span data-ttu-id="212a2-612">EWS 要求です。</span><span class="sxs-lookup"><span data-stu-id="212a2-612">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="212a2-613">関数</span><span class="sxs-lookup"><span data-stu-id="212a2-613">function</span></span>||<span data-ttu-id="212a2-614">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="212a2-614">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="212a2-p143">The XML result of the EWS call is provided as a string in the `asyncResult.value` property. If the result exceeds 1 MB in size, an error message is returned instead.</span><span class="sxs-lookup"><span data-stu-id="212a2-p143">The XML result of the EWS call is provided as a string in the `asyncResult.value` property. If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="212a2-617">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="212a2-617">Object</span></span>| <span data-ttu-id="212a2-618">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="212a2-618">&lt;optional&gt;</span></span>|<span data-ttu-id="212a2-619">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="212a2-619">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="212a2-620">要件</span><span class="sxs-lookup"><span data-stu-id="212a2-620">Requirements</span></span>

|<span data-ttu-id="212a2-621">要件</span><span class="sxs-lookup"><span data-stu-id="212a2-621">Requirement</span></span>| <span data-ttu-id="212a2-622">値</span><span class="sxs-lookup"><span data-stu-id="212a2-622">Value</span></span>|
|---|---|
|[<span data-ttu-id="212a2-623">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="212a2-623">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="212a2-624">1.0</span><span class="sxs-lookup"><span data-stu-id="212a2-624">1.0</span></span>|
|[<span data-ttu-id="212a2-625">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="212a2-625">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="212a2-626">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="212a2-626">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="212a2-627">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="212a2-627">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="212a2-628">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="212a2-628">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="212a2-629">例</span><span class="sxs-lookup"><span data-stu-id="212a2-629">Example</span></span>

<span data-ttu-id="212a2-630">次の例は、`GetItem` 操作を使ってアイテムの件名を取得するため、`makeEwsRequestAsync` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="212a2-630">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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

####  <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="212a2-631">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="212a2-631">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="212a2-632">サポートされているイベントの種類のイベント ハンドラーを削除します。</span><span class="sxs-lookup"><span data-stu-id="212a2-632">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="212a2-633">現在、サポートされている`Office.EventType.ItemChanged`イベント`Office.EventType.OfficeThemeChanged`の種類はおよびです。</span><span class="sxs-lookup"><span data-stu-id="212a2-633">Currently, the supported event types are `Office.EventType.ItemChanged` and `Office.EventType.OfficeThemeChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="212a2-634">パラメーター</span><span class="sxs-lookup"><span data-stu-id="212a2-634">Parameters</span></span>

| <span data-ttu-id="212a2-635">名前</span><span class="sxs-lookup"><span data-stu-id="212a2-635">Name</span></span> | <span data-ttu-id="212a2-636">型</span><span class="sxs-lookup"><span data-stu-id="212a2-636">Type</span></span> | <span data-ttu-id="212a2-637">属性</span><span class="sxs-lookup"><span data-stu-id="212a2-637">Attributes</span></span> | <span data-ttu-id="212a2-638">説明</span><span class="sxs-lookup"><span data-stu-id="212a2-638">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="212a2-639">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="212a2-639">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="212a2-640">ハンドラーを取り消すイベント。</span><span class="sxs-lookup"><span data-stu-id="212a2-640">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="212a2-641">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="212a2-641">Object</span></span> | <span data-ttu-id="212a2-642">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="212a2-642">&lt;optional&gt;</span></span> | <span data-ttu-id="212a2-643">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="212a2-643">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="212a2-644">Object</span><span class="sxs-lookup"><span data-stu-id="212a2-644">Object</span></span> | <span data-ttu-id="212a2-645">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="212a2-645">&lt;optional&gt;</span></span> | <span data-ttu-id="212a2-646">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="212a2-646">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="212a2-647">function</span><span class="sxs-lookup"><span data-stu-id="212a2-647">function</span></span>| <span data-ttu-id="212a2-648">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="212a2-648">&lt;optional&gt;</span></span>|<span data-ttu-id="212a2-649">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="212a2-649">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="212a2-650">要件</span><span class="sxs-lookup"><span data-stu-id="212a2-650">Requirements</span></span>

|<span data-ttu-id="212a2-651">要件</span><span class="sxs-lookup"><span data-stu-id="212a2-651">Requirement</span></span>| <span data-ttu-id="212a2-652">値</span><span class="sxs-lookup"><span data-stu-id="212a2-652">Value</span></span>|
|---|---|
|[<span data-ttu-id="212a2-653">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="212a2-653">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="212a2-654">1.5</span><span class="sxs-lookup"><span data-stu-id="212a2-654">1.5</span></span> |
|[<span data-ttu-id="212a2-655">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="212a2-655">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="212a2-656">ReadItem</span><span class="sxs-lookup"><span data-stu-id="212a2-656">ReadItem</span></span> |
|[<span data-ttu-id="212a2-657">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="212a2-657">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="212a2-658">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="212a2-658">Compose or Read</span></span>|
