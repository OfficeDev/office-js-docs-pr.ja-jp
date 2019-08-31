---
title: Office. メールボックス要件セット1.4
description: ''
ms.date: 08/30/2019
localization_priority: Normal
ms.openlocfilehash: 66ae7cb05ac56224fd7461c5c29587e21a24020a
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696212"
---
# <a name="mailbox"></a><span data-ttu-id="eb7e9-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="eb7e9-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="eb7e9-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="eb7e9-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="eb7e9-104">Microsoft Outlook の Outlook アドインオブジェクトモデルへのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="eb7e9-105">要件</span><span class="sxs-lookup"><span data-stu-id="eb7e9-105">Requirements</span></span>

|<span data-ttu-id="eb7e9-106">要件</span><span class="sxs-lookup"><span data-stu-id="eb7e9-106">Requirement</span></span>| <span data-ttu-id="eb7e9-107">値</span><span class="sxs-lookup"><span data-stu-id="eb7e9-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="eb7e9-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="eb7e9-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="eb7e9-109">1.0</span><span class="sxs-lookup"><span data-stu-id="eb7e9-109">1.0</span></span>|
|[<span data-ttu-id="eb7e9-110">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="eb7e9-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="eb7e9-111">制限あり</span><span class="sxs-lookup"><span data-stu-id="eb7e9-111">Restricted</span></span>|
|[<span data-ttu-id="eb7e9-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="eb7e9-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="eb7e9-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="eb7e9-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="eb7e9-114">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="eb7e9-114">Members and methods</span></span>

| <span data-ttu-id="eb7e9-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="eb7e9-115">Member</span></span> | <span data-ttu-id="eb7e9-116">種類</span><span class="sxs-lookup"><span data-stu-id="eb7e9-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="eb7e9-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="eb7e9-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="eb7e9-118">メンバー</span><span class="sxs-lookup"><span data-stu-id="eb7e9-118">Member</span></span> |
| [<span data-ttu-id="eb7e9-119">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="eb7e9-119">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="eb7e9-120">メソッド</span><span class="sxs-lookup"><span data-stu-id="eb7e9-120">Method</span></span> |
| [<span data-ttu-id="eb7e9-121">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="eb7e9-121">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="eb7e9-122">メソッド</span><span class="sxs-lookup"><span data-stu-id="eb7e9-122">Method</span></span> |
| [<span data-ttu-id="eb7e9-123">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="eb7e9-123">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="eb7e9-124">メソッド</span><span class="sxs-lookup"><span data-stu-id="eb7e9-124">Method</span></span> |
| [<span data-ttu-id="eb7e9-125">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="eb7e9-125">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="eb7e9-126">メソッド</span><span class="sxs-lookup"><span data-stu-id="eb7e9-126">Method</span></span> |
| [<span data-ttu-id="eb7e9-127">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="eb7e9-127">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="eb7e9-128">メソッド</span><span class="sxs-lookup"><span data-stu-id="eb7e9-128">Method</span></span> |
| [<span data-ttu-id="eb7e9-129">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="eb7e9-129">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="eb7e9-130">メソッド</span><span class="sxs-lookup"><span data-stu-id="eb7e9-130">Method</span></span> |
| [<span data-ttu-id="eb7e9-131">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="eb7e9-131">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="eb7e9-132">メソッド</span><span class="sxs-lookup"><span data-stu-id="eb7e9-132">Method</span></span> |
| [<span data-ttu-id="eb7e9-133">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="eb7e9-133">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="eb7e9-134">メソッド</span><span class="sxs-lookup"><span data-stu-id="eb7e9-134">Method</span></span> |
| [<span data-ttu-id="eb7e9-135">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="eb7e9-135">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="eb7e9-136">メソッド</span><span class="sxs-lookup"><span data-stu-id="eb7e9-136">Method</span></span> |
| [<span data-ttu-id="eb7e9-137">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="eb7e9-137">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="eb7e9-138">メソッド</span><span class="sxs-lookup"><span data-stu-id="eb7e9-138">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="eb7e9-139">名前空間</span><span class="sxs-lookup"><span data-stu-id="eb7e9-139">Namespaces</span></span>

<span data-ttu-id="eb7e9-140">[diagnostics](Office.context.mailbox.diagnostics.md):Outlook アドインに診断情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-140">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="eb7e9-141">[item](Office.context.mailbox.item.md):Outlook アドインのメッセージや予定にアクセスするためのメソッドとプロパティを提供します。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-141">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="eb7e9-142">[userProfile](Office.context.mailbox.userProfile.md):Outlook アドインのユーザーに関する情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-142">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="eb7e9-143">メンバー</span><span class="sxs-lookup"><span data-stu-id="eb7e9-143">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="eb7e9-144">ewsUrl: String</span><span class="sxs-lookup"><span data-stu-id="eb7e9-144">ewsUrl: String</span></span>

<span data-ttu-id="eb7e9-145">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account.</span><span class="sxs-lookup"><span data-stu-id="eb7e9-145">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account.</span></span> <span data-ttu-id="eb7e9-146">Read mode only.</span><span class="sxs-lookup"><span data-stu-id="eb7e9-146">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="eb7e9-147">このメンバーは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-147">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="eb7e9-p102">
  `ewsUrl\` 値は、リモート サービスで、ユーザーのメールボックスに EWS 呼び出しを行うために使うことができます。たとえば、[選択したアイテムから添付ファイルを取得する](/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="eb7e9-150">アプリが閲覧モードで `ewsUrl` メンバーを呼び出すには、アプリのマニフェスト内に **ReadItem** アクセス許可が指定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-150">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="eb7e9-p103">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出してから、`ewsUrl` メンバーを使用する必要があります。アプリには、`saveAsync` メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="eb7e9-153">型</span><span class="sxs-lookup"><span data-stu-id="eb7e9-153">Type</span></span>

*   <span data-ttu-id="eb7e9-154">String</span><span class="sxs-lookup"><span data-stu-id="eb7e9-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="eb7e9-155">要件</span><span class="sxs-lookup"><span data-stu-id="eb7e9-155">Requirements</span></span>

|<span data-ttu-id="eb7e9-156">要件</span><span class="sxs-lookup"><span data-stu-id="eb7e9-156">Requirement</span></span>| <span data-ttu-id="eb7e9-157">値</span><span class="sxs-lookup"><span data-stu-id="eb7e9-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="eb7e9-158">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="eb7e9-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="eb7e9-159">1.0</span><span class="sxs-lookup"><span data-stu-id="eb7e9-159">1.0</span></span>|
|[<span data-ttu-id="eb7e9-160">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="eb7e9-160">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="eb7e9-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="eb7e9-161">ReadItem</span></span>|
|[<span data-ttu-id="eb7e9-162">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="eb7e9-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="eb7e9-163">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="eb7e9-163">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="eb7e9-164">メソッド</span><span class="sxs-lookup"><span data-stu-id="eb7e9-164">Methods</span></span>

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="eb7e9-165">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="eb7e9-165">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="eb7e9-166">REST 形式のアイテム ID を EWS 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-166">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="eb7e9-167">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-167">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="eb7e9-p104">REST API ([Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) や [Microsoft Graph](https://graph.microsoft.io/) など) で取得されたアイテム ID は、Exchange Web サービス (EWS) に使用される形式とは異なる形式を使用します。`convertToEwsId` メソッドは、REST 形式の ID を EWS 用の適切な形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-p104">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="eb7e9-170">パラメーター</span><span class="sxs-lookup"><span data-stu-id="eb7e9-170">Parameters</span></span>

|<span data-ttu-id="eb7e9-171">名前</span><span class="sxs-lookup"><span data-stu-id="eb7e9-171">Name</span></span>| <span data-ttu-id="eb7e9-172">種類</span><span class="sxs-lookup"><span data-stu-id="eb7e9-172">Type</span></span>| <span data-ttu-id="eb7e9-173">説明</span><span class="sxs-lookup"><span data-stu-id="eb7e9-173">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="eb7e9-174">String</span><span class="sxs-lookup"><span data-stu-id="eb7e9-174">String</span></span>|<span data-ttu-id="eb7e9-175">Outlook REST API 形式のアイテム ID</span><span class="sxs-lookup"><span data-stu-id="eb7e9-175">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="eb7e9-176">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="eb7e9-176">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.4)|<span data-ttu-id="eb7e9-177">アイテム ID の取得に使用された Outlook REST API のバージョンを示す値。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-177">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="eb7e9-178">要件</span><span class="sxs-lookup"><span data-stu-id="eb7e9-178">Requirements</span></span>

|<span data-ttu-id="eb7e9-179">要件</span><span class="sxs-lookup"><span data-stu-id="eb7e9-179">Requirement</span></span>| <span data-ttu-id="eb7e9-180">値</span><span class="sxs-lookup"><span data-stu-id="eb7e9-180">Value</span></span>|
|---|---|
|[<span data-ttu-id="eb7e9-181">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="eb7e9-181">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="eb7e9-182">1.3</span><span class="sxs-lookup"><span data-stu-id="eb7e9-182">1.3</span></span>|
|[<span data-ttu-id="eb7e9-183">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="eb7e9-183">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="eb7e9-184">制限あり</span><span class="sxs-lookup"><span data-stu-id="eb7e9-184">Restricted</span></span>|
|[<span data-ttu-id="eb7e9-185">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="eb7e9-185">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="eb7e9-186">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="eb7e9-186">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="eb7e9-187">戻り値:</span><span class="sxs-lookup"><span data-stu-id="eb7e9-187">Returns:</span></span>

<span data-ttu-id="eb7e9-188">型:String</span><span class="sxs-lookup"><span data-stu-id="eb7e9-188">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="eb7e9-189">例</span><span class="sxs-lookup"><span data-stu-id="eb7e9-189">Example</span></span>

```js
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-14"></a><span data-ttu-id="eb7e9-190">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.4)}</span><span class="sxs-lookup"><span data-stu-id="eb7e9-190">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.4)}</span></span>

<span data-ttu-id="eb7e9-191">クライアントのローカル時間で時間情報が含まれている辞書を取得します。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-191">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="eb7e9-192">デスクトップまたは web 上の Outlook 用メールアプリは、日付と時刻に異なるタイムゾーンを使用できます。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-192">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times.</span></span> <span data-ttu-id="eb7e9-193">デスクトップ上の Outlook では、クライアントコンピューターのタイムゾーンが使用されます。Outlook on the web では、Exchange 管理センター (EAC) で設定されているタイムゾーンが使用されます。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-193">Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC).</span></span> <span data-ttu-id="eb7e9-194">日付と時刻の値を処理して、ユーザーインターフェイスに表示される値が、ユーザーが期待するタイムゾーンに常に一致するようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-194">You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="eb7e9-195">デスクトップクライアント上の Outlook でメールアプリが実行されている`convertToLocalClientTime`場合、このメソッドは、クライアントコンピューターのタイムゾーンに設定された値を持つ dictionary オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-195">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone.</span></span> <span data-ttu-id="eb7e9-196">メールアプリが web 上の Outlook で実行されている`convertToLocalClientTime`場合、このメソッドは、EAC で指定されたタイムゾーンに設定された値を持つ dictionary オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-196">If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="eb7e9-197">パラメーター</span><span class="sxs-lookup"><span data-stu-id="eb7e9-197">Parameters</span></span>

|<span data-ttu-id="eb7e9-198">名前</span><span class="sxs-lookup"><span data-stu-id="eb7e9-198">Name</span></span>| <span data-ttu-id="eb7e9-199">種類</span><span class="sxs-lookup"><span data-stu-id="eb7e9-199">Type</span></span>| <span data-ttu-id="eb7e9-200">説明</span><span class="sxs-lookup"><span data-stu-id="eb7e9-200">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="eb7e9-201">日付</span><span class="sxs-lookup"><span data-stu-id="eb7e9-201">Date</span></span>|<span data-ttu-id="eb7e9-202">日付オブジェクト</span><span class="sxs-lookup"><span data-stu-id="eb7e9-202">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="eb7e9-203">要件</span><span class="sxs-lookup"><span data-stu-id="eb7e9-203">Requirements</span></span>

|<span data-ttu-id="eb7e9-204">要件</span><span class="sxs-lookup"><span data-stu-id="eb7e9-204">Requirement</span></span>| <span data-ttu-id="eb7e9-205">値</span><span class="sxs-lookup"><span data-stu-id="eb7e9-205">Value</span></span>|
|---|---|
|[<span data-ttu-id="eb7e9-206">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="eb7e9-206">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="eb7e9-207">1.0</span><span class="sxs-lookup"><span data-stu-id="eb7e9-207">1.0</span></span>|
|[<span data-ttu-id="eb7e9-208">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="eb7e9-208">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="eb7e9-209">ReadItem</span><span class="sxs-lookup"><span data-stu-id="eb7e9-209">ReadItem</span></span>|
|[<span data-ttu-id="eb7e9-210">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="eb7e9-210">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="eb7e9-211">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="eb7e9-211">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="eb7e9-212">戻り値:</span><span class="sxs-lookup"><span data-stu-id="eb7e9-212">Returns:</span></span>

<span data-ttu-id="eb7e9-213">型:[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="eb7e9-213">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.4)</span></span>

<br>

---
---

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="eb7e9-214">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="eb7e9-214">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="eb7e9-215">EWS 形式のアイテム ID を REST 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-215">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="eb7e9-216">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-216">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="eb7e9-p107">EWS または `itemId` プロパティで取得されるアイテム ID は、REST API ([Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) や [Microsoft Graph](https://graph.microsoft.io/) など) に使用される形式とは異なる形式を使用します。`convertToRestId` メソッドは、EWS 形式の ID を REST 用の適切な形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-p107">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="eb7e9-219">パラメーター</span><span class="sxs-lookup"><span data-stu-id="eb7e9-219">Parameters</span></span>

|<span data-ttu-id="eb7e9-220">名前</span><span class="sxs-lookup"><span data-stu-id="eb7e9-220">Name</span></span>| <span data-ttu-id="eb7e9-221">種類</span><span class="sxs-lookup"><span data-stu-id="eb7e9-221">Type</span></span>| <span data-ttu-id="eb7e9-222">説明</span><span class="sxs-lookup"><span data-stu-id="eb7e9-222">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="eb7e9-223">String</span><span class="sxs-lookup"><span data-stu-id="eb7e9-223">String</span></span>|<span data-ttu-id="eb7e9-224">Exchange Web サービス (EWS) 形式のアイテム ID</span><span class="sxs-lookup"><span data-stu-id="eb7e9-224">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="eb7e9-225">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="eb7e9-225">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.4)|<span data-ttu-id="eb7e9-226">変換後の ID を使用する Outlook REST API のバージョンを示す値。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-226">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="eb7e9-227">要件</span><span class="sxs-lookup"><span data-stu-id="eb7e9-227">Requirements</span></span>

|<span data-ttu-id="eb7e9-228">要件</span><span class="sxs-lookup"><span data-stu-id="eb7e9-228">Requirement</span></span>| <span data-ttu-id="eb7e9-229">値</span><span class="sxs-lookup"><span data-stu-id="eb7e9-229">Value</span></span>|
|---|---|
|[<span data-ttu-id="eb7e9-230">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="eb7e9-230">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="eb7e9-231">1.3</span><span class="sxs-lookup"><span data-stu-id="eb7e9-231">1.3</span></span>|
|[<span data-ttu-id="eb7e9-232">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="eb7e9-232">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="eb7e9-233">制限あり</span><span class="sxs-lookup"><span data-stu-id="eb7e9-233">Restricted</span></span>|
|[<span data-ttu-id="eb7e9-234">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="eb7e9-234">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="eb7e9-235">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="eb7e9-235">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="eb7e9-236">戻り値:</span><span class="sxs-lookup"><span data-stu-id="eb7e9-236">Returns:</span></span>

<span data-ttu-id="eb7e9-237">型:String</span><span class="sxs-lookup"><span data-stu-id="eb7e9-237">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="eb7e9-238">例</span><span class="sxs-lookup"><span data-stu-id="eb7e9-238">Example</span></span>

```js
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="eb7e9-239">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="eb7e9-239">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="eb7e9-240">時間情報が含まれているディクショナリから日付オブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-240">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="eb7e9-241">`convertToUtcClientTime` メソッドは、ローカルの日付と時刻を含むディクショナリを、ローカルの日付と時刻の正しい値を持つ日付オブジェクトに変換します。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-241">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="eb7e9-242">パラメーター</span><span class="sxs-lookup"><span data-stu-id="eb7e9-242">Parameters</span></span>

|<span data-ttu-id="eb7e9-243">名前</span><span class="sxs-lookup"><span data-stu-id="eb7e9-243">Name</span></span>| <span data-ttu-id="eb7e9-244">型</span><span class="sxs-lookup"><span data-stu-id="eb7e9-244">Type</span></span>| <span data-ttu-id="eb7e9-245">説明</span><span class="sxs-lookup"><span data-stu-id="eb7e9-245">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="eb7e9-246">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="eb7e9-246">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.4)|<span data-ttu-id="eb7e9-247">変換するローカル時刻の値。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-247">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="eb7e9-248">要件</span><span class="sxs-lookup"><span data-stu-id="eb7e9-248">Requirements</span></span>

|<span data-ttu-id="eb7e9-249">要件</span><span class="sxs-lookup"><span data-stu-id="eb7e9-249">Requirement</span></span>| <span data-ttu-id="eb7e9-250">値</span><span class="sxs-lookup"><span data-stu-id="eb7e9-250">Value</span></span>|
|---|---|
|[<span data-ttu-id="eb7e9-251">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="eb7e9-251">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="eb7e9-252">1.0</span><span class="sxs-lookup"><span data-stu-id="eb7e9-252">1.0</span></span>|
|[<span data-ttu-id="eb7e9-253">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="eb7e9-253">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="eb7e9-254">ReadItem</span><span class="sxs-lookup"><span data-stu-id="eb7e9-254">ReadItem</span></span>|
|[<span data-ttu-id="eb7e9-255">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="eb7e9-255">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="eb7e9-256">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="eb7e9-256">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="eb7e9-257">戻り値:</span><span class="sxs-lookup"><span data-stu-id="eb7e9-257">Returns:</span></span>

<span data-ttu-id="eb7e9-258">時間が UTC で表現された日付オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-258">A Date object with the time expressed in UTC.</span></span>

<span data-ttu-id="eb7e9-259">型: Date</span><span class="sxs-lookup"><span data-stu-id="eb7e9-259">Type: Date</span></span>

##### <a name="example"></a><span data-ttu-id="eb7e9-260">例</span><span class="sxs-lookup"><span data-stu-id="eb7e9-260">Example</span></span>

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

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="eb7e9-261">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="eb7e9-261">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="eb7e9-262">既存の予定を表示します。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-262">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="eb7e9-263">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-263">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="eb7e9-264">`displayAppointmentForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存の予定を開きます。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-264">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="eb7e9-265">Outlook on the Mac では、このメソッドを使用して、定期的なアイテムの一部ではない単一の予定を表示したり、定期的なアイテムのマスター予定を表示したりすることはできませんが、一連のインスタンスを表示することはできません。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-265">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series.</span></span> <span data-ttu-id="eb7e9-266">これは、Mac 上の Outlook では、定期的なアイテムのインスタンスのプロパティ (アイテム ID を含む) にアクセスできないためです。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-266">This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="eb7e9-267">Web 上の Outlook では、フォームの本文が 32 KB 以下の文字である場合にのみ、このメソッドは指定されたフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-267">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="eb7e9-268">指定のアイテム識別子が既存の予定を表していない場合は、クライアント コンピューターまたはデバイスで空のウィンドウが開き、エラー メッセージは返されません。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-268">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="eb7e9-269">パラメーター</span><span class="sxs-lookup"><span data-stu-id="eb7e9-269">Parameters</span></span>

|<span data-ttu-id="eb7e9-270">名前</span><span class="sxs-lookup"><span data-stu-id="eb7e9-270">Name</span></span>| <span data-ttu-id="eb7e9-271">型</span><span class="sxs-lookup"><span data-stu-id="eb7e9-271">Type</span></span>| <span data-ttu-id="eb7e9-272">説明</span><span class="sxs-lookup"><span data-stu-id="eb7e9-272">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="eb7e9-273">String</span><span class="sxs-lookup"><span data-stu-id="eb7e9-273">String</span></span>|<span data-ttu-id="eb7e9-274">既存の予定の Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-274">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="eb7e9-275">要件</span><span class="sxs-lookup"><span data-stu-id="eb7e9-275">Requirements</span></span>

|<span data-ttu-id="eb7e9-276">要件</span><span class="sxs-lookup"><span data-stu-id="eb7e9-276">Requirement</span></span>| <span data-ttu-id="eb7e9-277">値</span><span class="sxs-lookup"><span data-stu-id="eb7e9-277">Value</span></span>|
|---|---|
|[<span data-ttu-id="eb7e9-278">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="eb7e9-278">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="eb7e9-279">1.0</span><span class="sxs-lookup"><span data-stu-id="eb7e9-279">1.0</span></span>|
|[<span data-ttu-id="eb7e9-280">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="eb7e9-280">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="eb7e9-281">ReadItem</span><span class="sxs-lookup"><span data-stu-id="eb7e9-281">ReadItem</span></span>|
|[<span data-ttu-id="eb7e9-282">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="eb7e9-282">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="eb7e9-283">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="eb7e9-283">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="eb7e9-284">例</span><span class="sxs-lookup"><span data-stu-id="eb7e9-284">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

<br>

---
---

#### <a name="displaymessageformitemid"></a><span data-ttu-id="eb7e9-285">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="eb7e9-285">displayMessageForm(itemId)</span></span>

<span data-ttu-id="eb7e9-286">既存のメッセージを表示します。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-286">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="eb7e9-287">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-287">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="eb7e9-288">`displayMessageForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存のメッセージを開きます。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-288">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="eb7e9-289">Web 上の Outlook では、フォームの本文が 32 KB の文字数以下の場合にのみ、このメソッドは指定されたフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-289">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="eb7e9-290">指定のアイテム識別子が既存のメッセージを表していない場合は、クライアント コンピューターにはメッセージは表示されず、エラー メッセージも返されません。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-290">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="eb7e9-p109">予定を表す `itemId` を含む `displayMessageForm` を使用しないでください。既存の予定を表示するには、`displayAppointmentForm` メソッドを使用します。新しい予定を作成するフォームを表示するには、`displayNewAppointmentForm` メソッドを使用します。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-p109">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="eb7e9-293">パラメーター</span><span class="sxs-lookup"><span data-stu-id="eb7e9-293">Parameters</span></span>

|<span data-ttu-id="eb7e9-294">名前</span><span class="sxs-lookup"><span data-stu-id="eb7e9-294">Name</span></span>| <span data-ttu-id="eb7e9-295">型</span><span class="sxs-lookup"><span data-stu-id="eb7e9-295">Type</span></span>| <span data-ttu-id="eb7e9-296">説明</span><span class="sxs-lookup"><span data-stu-id="eb7e9-296">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="eb7e9-297">String</span><span class="sxs-lookup"><span data-stu-id="eb7e9-297">String</span></span>|<span data-ttu-id="eb7e9-298">既存のメッセージの Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-298">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="eb7e9-299">要件</span><span class="sxs-lookup"><span data-stu-id="eb7e9-299">Requirements</span></span>

|<span data-ttu-id="eb7e9-300">要件</span><span class="sxs-lookup"><span data-stu-id="eb7e9-300">Requirement</span></span>| <span data-ttu-id="eb7e9-301">値</span><span class="sxs-lookup"><span data-stu-id="eb7e9-301">Value</span></span>|
|---|---|
|[<span data-ttu-id="eb7e9-302">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="eb7e9-302">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="eb7e9-303">1.0</span><span class="sxs-lookup"><span data-stu-id="eb7e9-303">1.0</span></span>|
|[<span data-ttu-id="eb7e9-304">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="eb7e9-304">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="eb7e9-305">ReadItem</span><span class="sxs-lookup"><span data-stu-id="eb7e9-305">ReadItem</span></span>|
|[<span data-ttu-id="eb7e9-306">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="eb7e9-306">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="eb7e9-307">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="eb7e9-307">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="eb7e9-308">例</span><span class="sxs-lookup"><span data-stu-id="eb7e9-308">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

<br>

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="eb7e9-309">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="eb7e9-309">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="eb7e9-310">新しい予定を作成するためのフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-310">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="eb7e9-311">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-311">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="eb7e9-p110">`displayNewAppointmentForm` メソッドを使用すると、ユーザーが新しい予定または会議を作成できるフォームが開きます。パラメーターを指定すると、予定のフォーム フィールドにパラメーターの内容が自動的に設定されます。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-p110">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="eb7e9-314">Outlook on the web およびモバイルデバイスでは、このメソッドは常に出席者フィールドを含むフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-314">In Outlook on the web and mobile devices, this method always displays a form with an attendees field.</span></span> <span data-ttu-id="eb7e9-315">入力引数として出席者を指定しないと、このメソッドにより **[保存]** ボタンのあるフォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-315">If you do not specify any attendees as input arguments, the method displays a form with a **Save** button.</span></span> <span data-ttu-id="eb7e9-316">出席者を指定した場合には、フォームにその出席者と **[送信]** ボタンが表示されます。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-316">If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="eb7e9-p112">Outlook リッチ クライアントと Outlook RT で、`requiredAttendees`、`optionalAttendees`、または `resources` パラメーターに出席者またはリソースを指定し、このメソッドを実行すると、**[送信]** ボタンがある会議フォームが表示されます。受信者を指定せずにこのメソッドを実行すると、**[保存して閉じる]** ボタンがある予定フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-p112">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="eb7e9-319">パラメーターのいずれかが指定のサイズ制限を超える場合、または不明なパラメーター名が指定されている場合は、例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-319">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="eb7e9-320">パラメーター</span><span class="sxs-lookup"><span data-stu-id="eb7e9-320">Parameters</span></span>

|<span data-ttu-id="eb7e9-321">名前</span><span class="sxs-lookup"><span data-stu-id="eb7e9-321">Name</span></span>| <span data-ttu-id="eb7e9-322">型</span><span class="sxs-lookup"><span data-stu-id="eb7e9-322">Type</span></span>| <span data-ttu-id="eb7e9-323">説明</span><span class="sxs-lookup"><span data-stu-id="eb7e9-323">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="eb7e9-324">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="eb7e9-324">Object</span></span> | <span data-ttu-id="eb7e9-325">新しい予定を記述するパラメーターのディクショナリ。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-325">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="eb7e9-326">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)&gt;</span><span class="sxs-lookup"><span data-stu-id="eb7e9-326">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)&gt;</span></span> | <span data-ttu-id="eb7e9-p113">予定に必要な各出席者について、メール アドレスを含む文字列の配列、または `EmailAddressDetails` オブジェクトを含む配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-p113">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="eb7e9-329">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)&gt;</span><span class="sxs-lookup"><span data-stu-id="eb7e9-329">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)&gt;</span></span> | <span data-ttu-id="eb7e9-p114">予定の各任意出席者について、メール アドレスを含む文字列の配列、または `EmailAddressDetails` オブジェクトを含む配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-p114">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="eb7e9-332">日付</span><span class="sxs-lookup"><span data-stu-id="eb7e9-332">Date</span></span> | <span data-ttu-id="eb7e9-333">予定の開始日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-333">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="eb7e9-334">日付</span><span class="sxs-lookup"><span data-stu-id="eb7e9-334">Date</span></span> | <span data-ttu-id="eb7e9-335">予定の終了日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-335">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="eb7e9-336">String</span><span class="sxs-lookup"><span data-stu-id="eb7e9-336">String</span></span> | <span data-ttu-id="eb7e9-p115">予定の場所を含む文字列。文字列は最大 255 文字に制限されます。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-p115">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="eb7e9-339">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="eb7e9-339">Array.&lt;String&gt;</span></span> | <span data-ttu-id="eb7e9-p116">予定に必要なリソースを含む文字列の配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-p116">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="eb7e9-342">String</span><span class="sxs-lookup"><span data-stu-id="eb7e9-342">String</span></span> | <span data-ttu-id="eb7e9-p117">予定の件名を含む文字列です。文字列は最大 255 文字に制限されます。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-p117">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="eb7e9-345">String</span><span class="sxs-lookup"><span data-stu-id="eb7e9-345">String</span></span> | <span data-ttu-id="eb7e9-p118">予定の本文。本文の内容は、最大サイズが 32 KB に制限されます。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-p118">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="eb7e9-348">要件</span><span class="sxs-lookup"><span data-stu-id="eb7e9-348">Requirements</span></span>

|<span data-ttu-id="eb7e9-349">要件</span><span class="sxs-lookup"><span data-stu-id="eb7e9-349">Requirement</span></span>| <span data-ttu-id="eb7e9-350">値</span><span class="sxs-lookup"><span data-stu-id="eb7e9-350">Value</span></span>|
|---|---|
|[<span data-ttu-id="eb7e9-351">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="eb7e9-351">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="eb7e9-352">1.0</span><span class="sxs-lookup"><span data-stu-id="eb7e9-352">1.0</span></span>|
|[<span data-ttu-id="eb7e9-353">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="eb7e9-353">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="eb7e9-354">ReadItem</span><span class="sxs-lookup"><span data-stu-id="eb7e9-354">ReadItem</span></span>|
|[<span data-ttu-id="eb7e9-355">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="eb7e9-355">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="eb7e9-356">読み取り</span><span class="sxs-lookup"><span data-stu-id="eb7e9-356">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="eb7e9-357">例</span><span class="sxs-lookup"><span data-stu-id="eb7e9-357">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="eb7e9-358">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="eb7e9-358">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="eb7e9-359">Exchange Server から添付ファイルやアイテムを取得するために使うトークンを含む文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-359">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="eb7e9-p119">`getCallbackTokenAsync` メソッドは、ユーザーのメールボックスをホストする Exchange Server から不透明なトークンを取得する非同期の呼び出しを行います。コールバック トークンの有効期間は 5 分です。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-p119">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="eb7e9-p120">トークンと予定の識別子またはアイテムの識別子をサードパーティ システムに渡すことができます。サードパーティ システムは、トークンをベアラー承認トークンとして使用し、Exchange Web サービス (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) または [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) 操作を呼び出して、添付ファイルまたはアイテムを返します。たとえば、[選択したアイテムから添付ファイルを取得する](/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-p120">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="eb7e9-365">アプリが閲覧モードで `getCallbackTokenAsync` メソッドを呼び出すには、アプリのマニフェスト内に **ReadItem** アクセス許可が指定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-365">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="eb7e9-p121">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出してアイテムの識別子を `getCallbackTokenAsync` メソッドに渡す必要があります。アプリには、`saveAsync` メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-p121">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="eb7e9-368">パラメーター</span><span class="sxs-lookup"><span data-stu-id="eb7e9-368">Parameters</span></span>

|<span data-ttu-id="eb7e9-369">名前</span><span class="sxs-lookup"><span data-stu-id="eb7e9-369">Name</span></span>| <span data-ttu-id="eb7e9-370">型</span><span class="sxs-lookup"><span data-stu-id="eb7e9-370">Type</span></span>| <span data-ttu-id="eb7e9-371">属性</span><span class="sxs-lookup"><span data-stu-id="eb7e9-371">Attributes</span></span>| <span data-ttu-id="eb7e9-372">説明</span><span class="sxs-lookup"><span data-stu-id="eb7e9-372">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="eb7e9-373">function</span><span class="sxs-lookup"><span data-stu-id="eb7e9-373">function</span></span>||<span data-ttu-id="eb7e9-374">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-374">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="eb7e9-375">トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-375">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="eb7e9-376">エラーが発生した場合は`asyncResult.error` 、 `asyncResult.diagnostics`プロパティとプロパティによって追加情報が提供されることがあります。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-376">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="eb7e9-377">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="eb7e9-377">Object</span></span>| <span data-ttu-id="eb7e9-378">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="eb7e9-378">&lt;optional&gt;</span></span>|<span data-ttu-id="eb7e9-379">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-379">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="eb7e9-380">エラー</span><span class="sxs-lookup"><span data-stu-id="eb7e9-380">Errors</span></span>

|<span data-ttu-id="eb7e9-381">エラー コード</span><span class="sxs-lookup"><span data-stu-id="eb7e9-381">Error code</span></span>|<span data-ttu-id="eb7e9-382">説明</span><span class="sxs-lookup"><span data-stu-id="eb7e9-382">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="eb7e9-383">要求が失敗しました。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-383">The request has failed.</span></span> <span data-ttu-id="eb7e9-384">HTTP エラーコードについては、diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-384">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="eb7e9-385">Exchange サーバーがエラーを返しました。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-385">The Exchange server returned an error.</span></span> <span data-ttu-id="eb7e9-386">詳細については、「diagnostics オブジェクト」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-386">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="eb7e9-387">ユーザーがネットワークに接続されていません。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-387">The user is no longer connected to the network.</span></span> <span data-ttu-id="eb7e9-388">ネットワーク接続を確認し、もう一度実行してください。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-388">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="eb7e9-389">要件</span><span class="sxs-lookup"><span data-stu-id="eb7e9-389">Requirements</span></span>

|<span data-ttu-id="eb7e9-390">要件</span><span class="sxs-lookup"><span data-stu-id="eb7e9-390">Requirement</span></span>| <span data-ttu-id="eb7e9-391">値</span><span class="sxs-lookup"><span data-stu-id="eb7e9-391">Value</span></span>|
|---|---|
|[<span data-ttu-id="eb7e9-392">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="eb7e9-392">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="eb7e9-393">1.0</span><span class="sxs-lookup"><span data-stu-id="eb7e9-393">1.0</span></span>|
|[<span data-ttu-id="eb7e9-394">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="eb7e9-394">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="eb7e9-395">ReadItem</span><span class="sxs-lookup"><span data-stu-id="eb7e9-395">ReadItem</span></span>|
|[<span data-ttu-id="eb7e9-396">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="eb7e9-396">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="eb7e9-397">新規作成と閲覧</span><span class="sxs-lookup"><span data-stu-id="eb7e9-397">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="eb7e9-398">例</span><span class="sxs-lookup"><span data-stu-id="eb7e9-398">Example</span></span>

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

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="eb7e9-399">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="eb7e9-399">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="eb7e9-400">ユーザーと Office アドインを識別するトークンを取得します。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-400">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="eb7e9-401">`getUserIdentityTokenAsync` メソッドは、[アドインとユーザーをサード パーティのシステムで識別して認証](/outlook/add-ins/authentication)することのできるトークンを返します。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-401">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="eb7e9-402">パラメーター</span><span class="sxs-lookup"><span data-stu-id="eb7e9-402">Parameters</span></span>

|<span data-ttu-id="eb7e9-403">名前</span><span class="sxs-lookup"><span data-stu-id="eb7e9-403">Name</span></span>| <span data-ttu-id="eb7e9-404">型</span><span class="sxs-lookup"><span data-stu-id="eb7e9-404">Type</span></span>| <span data-ttu-id="eb7e9-405">属性</span><span class="sxs-lookup"><span data-stu-id="eb7e9-405">Attributes</span></span>| <span data-ttu-id="eb7e9-406">説明</span><span class="sxs-lookup"><span data-stu-id="eb7e9-406">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="eb7e9-407">function</span><span class="sxs-lookup"><span data-stu-id="eb7e9-407">function</span></span>||<span data-ttu-id="eb7e9-408">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-408">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="eb7e9-409">トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-409">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="eb7e9-410">エラーが発生した場合は`asyncResult.error` 、 `asyncResult.diagnostics`プロパティとプロパティによって追加情報が提供されることがあります。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-410">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="eb7e9-411">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="eb7e9-411">Object</span></span>| <span data-ttu-id="eb7e9-412">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="eb7e9-412">&lt;optional&gt;</span></span>|<span data-ttu-id="eb7e9-413">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-413">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="eb7e9-414">エラー</span><span class="sxs-lookup"><span data-stu-id="eb7e9-414">Errors</span></span>

|<span data-ttu-id="eb7e9-415">エラー コード</span><span class="sxs-lookup"><span data-stu-id="eb7e9-415">Error code</span></span>|<span data-ttu-id="eb7e9-416">説明</span><span class="sxs-lookup"><span data-stu-id="eb7e9-416">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="eb7e9-417">要求が失敗しました。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-417">The request has failed.</span></span> <span data-ttu-id="eb7e9-418">HTTP エラーコードについては、diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-418">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="eb7e9-419">Exchange サーバーがエラーを返しました。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-419">The Exchange server returned an error.</span></span> <span data-ttu-id="eb7e9-420">詳細については、「diagnostics オブジェクト」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-420">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="eb7e9-421">ユーザーがネットワークに接続されていません。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-421">The user is no longer connected to the network.</span></span> <span data-ttu-id="eb7e9-422">ネットワーク接続を確認し、もう一度実行してください。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-422">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="eb7e9-423">要件</span><span class="sxs-lookup"><span data-stu-id="eb7e9-423">Requirements</span></span>

|<span data-ttu-id="eb7e9-424">要件</span><span class="sxs-lookup"><span data-stu-id="eb7e9-424">Requirement</span></span>| <span data-ttu-id="eb7e9-425">値</span><span class="sxs-lookup"><span data-stu-id="eb7e9-425">Value</span></span>|
|---|---|
|[<span data-ttu-id="eb7e9-426">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="eb7e9-426">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="eb7e9-427">1.0</span><span class="sxs-lookup"><span data-stu-id="eb7e9-427">1.0</span></span>|
|[<span data-ttu-id="eb7e9-428">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="eb7e9-428">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="eb7e9-429">ReadItem</span><span class="sxs-lookup"><span data-stu-id="eb7e9-429">ReadItem</span></span>|
|[<span data-ttu-id="eb7e9-430">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="eb7e9-430">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="eb7e9-431">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="eb7e9-431">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="eb7e9-432">例</span><span class="sxs-lookup"><span data-stu-id="eb7e9-432">Example</span></span>

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

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="eb7e9-433">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="eb7e9-433">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="eb7e9-434">ユーザーのメールボックスをホストしている Exchange サーバー上の Exchange Web サービス (EWS) のサービスに対して非同期の要求を行います。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-434">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="eb7e9-435">このメソッドは、次のシナリオではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-435">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="eb7e9-436">Outlook on iOS または Android</span><span class="sxs-lookup"><span data-stu-id="eb7e9-436">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="eb7e9-437">アドインが Gmail のメールボックスに読み込まれる場合</span><span class="sxs-lookup"><span data-stu-id="eb7e9-437">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="eb7e9-438">このような場合は、アドインでは [REST API を使用](/outlook/add-ins/use-rest-api)して、代わりにユーザーのメールボックスにアクセスする必要があります。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-438">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="eb7e9-p128">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange. See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span><span class="sxs-lookup"><span data-stu-id="eb7e9-p128">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange. See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="eb7e9-441">`makeEwsRequestAsync` メソッドでは、フォルダー関連アイテムを要求できません。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-441">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="eb7e9-442">XML 要求では UTF-8 エンコードを指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-442">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="eb7e9-p129">`makeEwsRequestAsync` メソッドを使用するには、アドインに **ReadWriteMailbox** アクセス許可が必要です。**ReadWriteMailbox** アクセス許可と、`makeEwsRequestAsync` メソッドで呼び出せる EWS 操作の使い方については、「[ユーザーのメールボックスへのメール アドイン アクセスのアクセス許可を指定する](/outlook/add-ins/understanding-outlook-add-in-permissions)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-p129">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="eb7e9-445">サーバー管理者は、クライアント アクセス サーバーの EWS ディレクトリで `OAuthAuthentication` を true に設定して、`makeEwsRequestAsync` メソッドで EWS 要求を行うことができるようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-445">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="eb7e9-446">バージョンの相違点</span><span class="sxs-lookup"><span data-stu-id="eb7e9-446">Version differences</span></span>

<span data-ttu-id="eb7e9-447">バージョン 15.0.4535.1004 より前のバージョンの Outlook で実行しているメール アプリで `makeEwsRequestAsync` メソッドを使う場合は、エンコード値を `ISO-8859-1` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-447">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="eb7e9-p130">Outlook on the web でメール アプリを実行している場合は、エンコード値を設定する必要はありません。mailbox.diagnostics.hostName プロパティを使って、メール アプリを Outlook で実行しているのか、Outlook on the web で実行しているのかを確認できます。mailbox.diagnostics.hostVersion プロパティを使って、どのバージョンの Outlook を使って実行しているかを確認できます。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-p130">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="eb7e9-451">パラメーター</span><span class="sxs-lookup"><span data-stu-id="eb7e9-451">Parameters</span></span>

|<span data-ttu-id="eb7e9-452">名前</span><span class="sxs-lookup"><span data-stu-id="eb7e9-452">Name</span></span>| <span data-ttu-id="eb7e9-453">型</span><span class="sxs-lookup"><span data-stu-id="eb7e9-453">Type</span></span>| <span data-ttu-id="eb7e9-454">属性</span><span class="sxs-lookup"><span data-stu-id="eb7e9-454">Attributes</span></span>| <span data-ttu-id="eb7e9-455">説明</span><span class="sxs-lookup"><span data-stu-id="eb7e9-455">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="eb7e9-456">String</span><span class="sxs-lookup"><span data-stu-id="eb7e9-456">String</span></span>||<span data-ttu-id="eb7e9-457">EWS 要求です。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-457">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="eb7e9-458">function</span><span class="sxs-lookup"><span data-stu-id="eb7e9-458">function</span></span>||<span data-ttu-id="eb7e9-459">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-459">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="eb7e9-p131">The XML result of the EWS call is provided as a string in the `asyncResult.value` property. If the result exceeds 1 MB in size, an error message is returned instead.</span><span class="sxs-lookup"><span data-stu-id="eb7e9-p131">The XML result of the EWS call is provided as a string in the `asyncResult.value` property. If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="eb7e9-462">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="eb7e9-462">Object</span></span>| <span data-ttu-id="eb7e9-463">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="eb7e9-463">&lt;optional&gt;</span></span>|<span data-ttu-id="eb7e9-464">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-464">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="eb7e9-465">要件</span><span class="sxs-lookup"><span data-stu-id="eb7e9-465">Requirements</span></span>

|<span data-ttu-id="eb7e9-466">要件</span><span class="sxs-lookup"><span data-stu-id="eb7e9-466">Requirement</span></span>| <span data-ttu-id="eb7e9-467">値</span><span class="sxs-lookup"><span data-stu-id="eb7e9-467">Value</span></span>|
|---|---|
|[<span data-ttu-id="eb7e9-468">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="eb7e9-468">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="eb7e9-469">1.0</span><span class="sxs-lookup"><span data-stu-id="eb7e9-469">1.0</span></span>|
|[<span data-ttu-id="eb7e9-470">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="eb7e9-470">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="eb7e9-471">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="eb7e9-471">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="eb7e9-472">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="eb7e9-472">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="eb7e9-473">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="eb7e9-473">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="eb7e9-474">例</span><span class="sxs-lookup"><span data-stu-id="eb7e9-474">Example</span></span>

<span data-ttu-id="eb7e9-475">次の例は、`makeEwsRequestAsync` を呼び出し、`GetItem` 操作を使って項目の件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="eb7e9-475">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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
