---
title: Office. メールボックス要件セット1.4
description: ''
ms.date: 10/21/2019
localization_priority: Normal
ms.openlocfilehash: 46a73e4911d95310efbe0607b6ba0715238cd6cc
ms.sourcegitcommit: 499bf49b41205f8034c501d4db5fe4b02dab205e
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/22/2019
ms.locfileid: "37626986"
---
# <a name="mailbox"></a><span data-ttu-id="5179a-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="5179a-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="5179a-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="5179a-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="5179a-104">Microsoft Outlook の Outlook アドイン オブジェクト モデルへのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="5179a-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="5179a-105">要件</span><span class="sxs-lookup"><span data-stu-id="5179a-105">Requirements</span></span>

|<span data-ttu-id="5179a-106">要件</span><span class="sxs-lookup"><span data-stu-id="5179a-106">Requirement</span></span>| <span data-ttu-id="5179a-107">値</span><span class="sxs-lookup"><span data-stu-id="5179a-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="5179a-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="5179a-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5179a-109">1.0</span><span class="sxs-lookup"><span data-stu-id="5179a-109">1.0</span></span>|
|[<span data-ttu-id="5179a-110">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="5179a-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5179a-111">制限あり</span><span class="sxs-lookup"><span data-stu-id="5179a-111">Restricted</span></span>|
|[<span data-ttu-id="5179a-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="5179a-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5179a-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="5179a-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="5179a-114">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="5179a-114">Members and methods</span></span>

| <span data-ttu-id="5179a-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="5179a-115">Member</span></span> | <span data-ttu-id="5179a-116">種類</span><span class="sxs-lookup"><span data-stu-id="5179a-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="5179a-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="5179a-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="5179a-118">メンバー</span><span class="sxs-lookup"><span data-stu-id="5179a-118">Member</span></span> |
| [<span data-ttu-id="5179a-119">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="5179a-119">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="5179a-120">メソッド</span><span class="sxs-lookup"><span data-stu-id="5179a-120">Method</span></span> |
| [<span data-ttu-id="5179a-121">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="5179a-121">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="5179a-122">メソッド</span><span class="sxs-lookup"><span data-stu-id="5179a-122">Method</span></span> |
| [<span data-ttu-id="5179a-123">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="5179a-123">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="5179a-124">メソッド</span><span class="sxs-lookup"><span data-stu-id="5179a-124">Method</span></span> |
| [<span data-ttu-id="5179a-125">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="5179a-125">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="5179a-126">メソッド</span><span class="sxs-lookup"><span data-stu-id="5179a-126">Method</span></span> |
| [<span data-ttu-id="5179a-127">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="5179a-127">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="5179a-128">メソッド</span><span class="sxs-lookup"><span data-stu-id="5179a-128">Method</span></span> |
| [<span data-ttu-id="5179a-129">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="5179a-129">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="5179a-130">メソッド</span><span class="sxs-lookup"><span data-stu-id="5179a-130">Method</span></span> |
| [<span data-ttu-id="5179a-131">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="5179a-131">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="5179a-132">メソッド</span><span class="sxs-lookup"><span data-stu-id="5179a-132">Method</span></span> |
| [<span data-ttu-id="5179a-133">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="5179a-133">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="5179a-134">メソッド</span><span class="sxs-lookup"><span data-stu-id="5179a-134">Method</span></span> |
| [<span data-ttu-id="5179a-135">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="5179a-135">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="5179a-136">メソッド</span><span class="sxs-lookup"><span data-stu-id="5179a-136">Method</span></span> |
| [<span data-ttu-id="5179a-137">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="5179a-137">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="5179a-138">メソッド</span><span class="sxs-lookup"><span data-stu-id="5179a-138">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="5179a-139">名前空間</span><span class="sxs-lookup"><span data-stu-id="5179a-139">Namespaces</span></span>

<span data-ttu-id="5179a-140">[diagnostics](Office.context.mailbox.diagnostics.md):Outlook アドインに診断情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="5179a-140">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="5179a-141">[item](Office.context.mailbox.item.md):Outlook アドインのメッセージや予定にアクセスするためのメソッドとプロパティを提供します。</span><span class="sxs-lookup"><span data-stu-id="5179a-141">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="5179a-142">[userProfile](Office.context.mailbox.userProfile.md):Outlook アドインのユーザーに関する情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="5179a-142">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="5179a-143">Members</span><span class="sxs-lookup"><span data-stu-id="5179a-143">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="5179a-144">ewsUrl: String</span><span class="sxs-lookup"><span data-stu-id="5179a-144">ewsUrl: String</span></span>

<span data-ttu-id="5179a-p101">このメール アカウントの Exchange Web サービス (EWS) エンドポイントの URL を取得します。読み取りモードのみです。</span><span class="sxs-lookup"><span data-stu-id="5179a-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="5179a-147">このメンバーは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="5179a-147">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="5179a-p102">
  `ewsUrl\` 値は、リモート サービスで、ユーザーのメールボックスに EWS 呼び出しを行うために使うことができます。たとえば、[選択したアイテムから添付ファイルを取得する](/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="5179a-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="5179a-150">アプリが閲覧モードで `ewsUrl` メンバーを呼び出すには、アプリのマニフェスト内に **ReadItem** アクセス許可が指定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="5179a-150">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="5179a-p103">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出してから、`ewsUrl` メンバーを使用する必要があります。アプリには、`saveAsync` メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="5179a-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="5179a-153">型</span><span class="sxs-lookup"><span data-stu-id="5179a-153">Type</span></span>

*   <span data-ttu-id="5179a-154">String</span><span class="sxs-lookup"><span data-stu-id="5179a-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5179a-155">要件</span><span class="sxs-lookup"><span data-stu-id="5179a-155">Requirements</span></span>

|<span data-ttu-id="5179a-156">要件</span><span class="sxs-lookup"><span data-stu-id="5179a-156">Requirement</span></span>| <span data-ttu-id="5179a-157">値</span><span class="sxs-lookup"><span data-stu-id="5179a-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="5179a-158">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="5179a-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5179a-159">1.0</span><span class="sxs-lookup"><span data-stu-id="5179a-159">1.0</span></span>|
|[<span data-ttu-id="5179a-160">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="5179a-160">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5179a-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5179a-161">ReadItem</span></span>|
|[<span data-ttu-id="5179a-162">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="5179a-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5179a-163">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="5179a-163">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="5179a-164">メソッド</span><span class="sxs-lookup"><span data-stu-id="5179a-164">Methods</span></span>

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="5179a-165">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="5179a-165">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="5179a-166">REST 形式のアイテム ID を EWS 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="5179a-166">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="5179a-167">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="5179a-167">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="5179a-p104">REST API ([Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) や [Microsoft Graph](https://graph.microsoft.io/) など) で取得されたアイテム ID は、Exchange Web サービス (EWS) に使用される形式とは異なる形式を使用します。`convertToEwsId` メソッドは、REST 形式の ID を EWS 用の適切な形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="5179a-p104">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5179a-170">パラメーター</span><span class="sxs-lookup"><span data-stu-id="5179a-170">Parameters</span></span>

|<span data-ttu-id="5179a-171">名前</span><span class="sxs-lookup"><span data-stu-id="5179a-171">Name</span></span>| <span data-ttu-id="5179a-172">種類</span><span class="sxs-lookup"><span data-stu-id="5179a-172">Type</span></span>| <span data-ttu-id="5179a-173">説明</span><span class="sxs-lookup"><span data-stu-id="5179a-173">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="5179a-174">String</span><span class="sxs-lookup"><span data-stu-id="5179a-174">String</span></span>|<span data-ttu-id="5179a-175">Outlook REST API 形式のアイテム ID</span><span class="sxs-lookup"><span data-stu-id="5179a-175">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="5179a-176">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="5179a-176">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.4)|<span data-ttu-id="5179a-177">アイテム ID の取得に使用された Outlook REST API のバージョンを示す値。</span><span class="sxs-lookup"><span data-stu-id="5179a-177">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5179a-178">要件</span><span class="sxs-lookup"><span data-stu-id="5179a-178">Requirements</span></span>

|<span data-ttu-id="5179a-179">要件</span><span class="sxs-lookup"><span data-stu-id="5179a-179">Requirement</span></span>| <span data-ttu-id="5179a-180">値</span><span class="sxs-lookup"><span data-stu-id="5179a-180">Value</span></span>|
|---|---|
|[<span data-ttu-id="5179a-181">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="5179a-181">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5179a-182">1.3</span><span class="sxs-lookup"><span data-stu-id="5179a-182">1.3</span></span>|
|[<span data-ttu-id="5179a-183">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="5179a-183">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5179a-184">制限あり</span><span class="sxs-lookup"><span data-stu-id="5179a-184">Restricted</span></span>|
|[<span data-ttu-id="5179a-185">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="5179a-185">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5179a-186">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="5179a-186">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="5179a-187">戻り値:</span><span class="sxs-lookup"><span data-stu-id="5179a-187">Returns:</span></span>

<span data-ttu-id="5179a-188">型:String</span><span class="sxs-lookup"><span data-stu-id="5179a-188">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="5179a-189">例</span><span class="sxs-lookup"><span data-stu-id="5179a-189">Example</span></span>

```js
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-14"></a><span data-ttu-id="5179a-190">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.4)}</span><span class="sxs-lookup"><span data-stu-id="5179a-190">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.4)}</span></span>

<span data-ttu-id="5179a-191">クライアントのローカル時間で時間情報が含まれている辞書を取得します。</span><span class="sxs-lookup"><span data-stu-id="5179a-191">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="5179a-p105">Outlook on the web または Outlook デスクトップのメールアプリでは、日付と時刻に異なるタイムゾーンを使用できます。Outlook デスクトップは、クライアント コンピューターのタイム ゾーンを使用します。Outlook on the web は、Exchange 管理センター (EAC) で設定されたタイム ゾーンを使用します。日付と時刻の値は、ユーザー インターフェイスに表示される値が、常にユーザーが期待するタイム ゾーンと一致するように処理する必要があります。</span><span class="sxs-lookup"><span data-stu-id="5179a-p105">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times. Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="5179a-p106">Outlook デスクトップ クライアントでメール アプリを実行している場合、`convertToLocalClientTime` メソッドは、クライアント コンピューターのタイム ゾーンに設定された値のディクショナリ オブジェクトを返します。Outlook on the web でメール アプリを実行している場合、`convertToLocalClientTime` メソッドは、EAC で指定したタイム ゾーンに設定された値のディクショナリ オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="5179a-p106">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5179a-197">パラメーター</span><span class="sxs-lookup"><span data-stu-id="5179a-197">Parameters</span></span>

|<span data-ttu-id="5179a-198">名前</span><span class="sxs-lookup"><span data-stu-id="5179a-198">Name</span></span>| <span data-ttu-id="5179a-199">種類</span><span class="sxs-lookup"><span data-stu-id="5179a-199">Type</span></span>| <span data-ttu-id="5179a-200">説明</span><span class="sxs-lookup"><span data-stu-id="5179a-200">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="5179a-201">日付</span><span class="sxs-lookup"><span data-stu-id="5179a-201">Date</span></span>|<span data-ttu-id="5179a-202">日付オブジェクト</span><span class="sxs-lookup"><span data-stu-id="5179a-202">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5179a-203">要件</span><span class="sxs-lookup"><span data-stu-id="5179a-203">Requirements</span></span>

|<span data-ttu-id="5179a-204">要件</span><span class="sxs-lookup"><span data-stu-id="5179a-204">Requirement</span></span>| <span data-ttu-id="5179a-205">値</span><span class="sxs-lookup"><span data-stu-id="5179a-205">Value</span></span>|
|---|---|
|[<span data-ttu-id="5179a-206">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="5179a-206">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5179a-207">1.0</span><span class="sxs-lookup"><span data-stu-id="5179a-207">1.0</span></span>|
|[<span data-ttu-id="5179a-208">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="5179a-208">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5179a-209">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5179a-209">ReadItem</span></span>|
|[<span data-ttu-id="5179a-210">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="5179a-210">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5179a-211">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="5179a-211">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="5179a-212">戻り値:</span><span class="sxs-lookup"><span data-stu-id="5179a-212">Returns:</span></span>

<span data-ttu-id="5179a-213">型:[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="5179a-213">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.4)</span></span>

<br>

---
---

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="5179a-214">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="5179a-214">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="5179a-215">EWS 形式のアイテム ID を REST 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="5179a-215">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="5179a-216">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="5179a-216">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="5179a-p107">EWS または `itemId` プロパティで取得されるアイテム ID は、REST API ([Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) や [Microsoft Graph](https://graph.microsoft.io/) など) に使用される形式とは異なる形式を使用します。`convertToRestId` メソッドは、EWS 形式の ID を REST 用の適切な形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="5179a-p107">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5179a-219">パラメーター</span><span class="sxs-lookup"><span data-stu-id="5179a-219">Parameters</span></span>

|<span data-ttu-id="5179a-220">名前</span><span class="sxs-lookup"><span data-stu-id="5179a-220">Name</span></span>| <span data-ttu-id="5179a-221">種類</span><span class="sxs-lookup"><span data-stu-id="5179a-221">Type</span></span>| <span data-ttu-id="5179a-222">説明</span><span class="sxs-lookup"><span data-stu-id="5179a-222">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="5179a-223">String</span><span class="sxs-lookup"><span data-stu-id="5179a-223">String</span></span>|<span data-ttu-id="5179a-224">Exchange Web サービス (EWS) 形式のアイテム ID</span><span class="sxs-lookup"><span data-stu-id="5179a-224">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="5179a-225">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="5179a-225">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.4)|<span data-ttu-id="5179a-226">変換後の ID を使用する Outlook REST API のバージョンを示す値。</span><span class="sxs-lookup"><span data-stu-id="5179a-226">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5179a-227">要件</span><span class="sxs-lookup"><span data-stu-id="5179a-227">Requirements</span></span>

|<span data-ttu-id="5179a-228">要件</span><span class="sxs-lookup"><span data-stu-id="5179a-228">Requirement</span></span>| <span data-ttu-id="5179a-229">値</span><span class="sxs-lookup"><span data-stu-id="5179a-229">Value</span></span>|
|---|---|
|[<span data-ttu-id="5179a-230">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="5179a-230">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5179a-231">1.3</span><span class="sxs-lookup"><span data-stu-id="5179a-231">1.3</span></span>|
|[<span data-ttu-id="5179a-232">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="5179a-232">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5179a-233">制限あり</span><span class="sxs-lookup"><span data-stu-id="5179a-233">Restricted</span></span>|
|[<span data-ttu-id="5179a-234">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="5179a-234">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5179a-235">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="5179a-235">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="5179a-236">戻り値:</span><span class="sxs-lookup"><span data-stu-id="5179a-236">Returns:</span></span>

<span data-ttu-id="5179a-237">型:String</span><span class="sxs-lookup"><span data-stu-id="5179a-237">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="5179a-238">例</span><span class="sxs-lookup"><span data-stu-id="5179a-238">Example</span></span>

```js
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="5179a-239">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="5179a-239">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="5179a-240">時間情報が含まれているディクショナリから日付オブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="5179a-240">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="5179a-241">`convertToUtcClientTime` メソッドは、ローカルの日付と時刻を含むディクショナリを、ローカルの日付と時刻の正しい値を持つ日付オブジェクトに変換します。</span><span class="sxs-lookup"><span data-stu-id="5179a-241">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5179a-242">パラメーター</span><span class="sxs-lookup"><span data-stu-id="5179a-242">Parameters</span></span>

|<span data-ttu-id="5179a-243">名前</span><span class="sxs-lookup"><span data-stu-id="5179a-243">Name</span></span>| <span data-ttu-id="5179a-244">種類</span><span class="sxs-lookup"><span data-stu-id="5179a-244">Type</span></span>| <span data-ttu-id="5179a-245">説明</span><span class="sxs-lookup"><span data-stu-id="5179a-245">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="5179a-246">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="5179a-246">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.4)|<span data-ttu-id="5179a-247">変換するローカル時刻の値。</span><span class="sxs-lookup"><span data-stu-id="5179a-247">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5179a-248">要件</span><span class="sxs-lookup"><span data-stu-id="5179a-248">Requirements</span></span>

|<span data-ttu-id="5179a-249">要件</span><span class="sxs-lookup"><span data-stu-id="5179a-249">Requirement</span></span>| <span data-ttu-id="5179a-250">値</span><span class="sxs-lookup"><span data-stu-id="5179a-250">Value</span></span>|
|---|---|
|[<span data-ttu-id="5179a-251">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="5179a-251">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5179a-252">1.0</span><span class="sxs-lookup"><span data-stu-id="5179a-252">1.0</span></span>|
|[<span data-ttu-id="5179a-253">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="5179a-253">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5179a-254">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5179a-254">ReadItem</span></span>|
|[<span data-ttu-id="5179a-255">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="5179a-255">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5179a-256">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="5179a-256">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="5179a-257">戻り値:</span><span class="sxs-lookup"><span data-stu-id="5179a-257">Returns:</span></span>

<span data-ttu-id="5179a-258">時間が UTC で表現された日付オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="5179a-258">A Date object with the time expressed in UTC.</span></span>

<span data-ttu-id="5179a-259">型: Date</span><span class="sxs-lookup"><span data-stu-id="5179a-259">Type: Date</span></span>

##### <a name="example"></a><span data-ttu-id="5179a-260">例</span><span class="sxs-lookup"><span data-stu-id="5179a-260">Example</span></span>

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

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="5179a-261">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="5179a-261">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="5179a-262">既存の予定を表示します。</span><span class="sxs-lookup"><span data-stu-id="5179a-262">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="5179a-263">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="5179a-263">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="5179a-264">`displayAppointmentForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存の予定を開きます。</span><span class="sxs-lookup"><span data-stu-id="5179a-264">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="5179a-p108">Outlook on Mac では、このメソッドを使用して定期的な系列に含まれない単発の予定や定期的な系列のマスター予定を表示できます。ただし、系列のインスタンスは表示できません。これは、Outlook on Mac では定期的な系列のインスタンスのプロパティ (アイテム ID を含む) にアクセスできないためです。</span><span class="sxs-lookup"><span data-stu-id="5179a-p108">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="5179a-267">Outlook on the web では、このメソッドはフォームの本文が 32KB 以下の文字数の場合にのみ指定のフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="5179a-267">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="5179a-268">指定のアイテム識別子が既存の予定を表していない場合は、クライアント コンピューターまたはデバイスで空のウィンドウが開き、エラー メッセージは返されません。</span><span class="sxs-lookup"><span data-stu-id="5179a-268">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5179a-269">パラメーター</span><span class="sxs-lookup"><span data-stu-id="5179a-269">Parameters</span></span>

|<span data-ttu-id="5179a-270">名前</span><span class="sxs-lookup"><span data-stu-id="5179a-270">Name</span></span>| <span data-ttu-id="5179a-271">種類</span><span class="sxs-lookup"><span data-stu-id="5179a-271">Type</span></span>| <span data-ttu-id="5179a-272">説明</span><span class="sxs-lookup"><span data-stu-id="5179a-272">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="5179a-273">String</span><span class="sxs-lookup"><span data-stu-id="5179a-273">String</span></span>|<span data-ttu-id="5179a-274">既存の予定の Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="5179a-274">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5179a-275">要件</span><span class="sxs-lookup"><span data-stu-id="5179a-275">Requirements</span></span>

|<span data-ttu-id="5179a-276">要件</span><span class="sxs-lookup"><span data-stu-id="5179a-276">Requirement</span></span>| <span data-ttu-id="5179a-277">値</span><span class="sxs-lookup"><span data-stu-id="5179a-277">Value</span></span>|
|---|---|
|[<span data-ttu-id="5179a-278">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="5179a-278">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5179a-279">1.0</span><span class="sxs-lookup"><span data-stu-id="5179a-279">1.0</span></span>|
|[<span data-ttu-id="5179a-280">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="5179a-280">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5179a-281">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5179a-281">ReadItem</span></span>|
|[<span data-ttu-id="5179a-282">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="5179a-282">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5179a-283">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="5179a-283">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5179a-284">例</span><span class="sxs-lookup"><span data-stu-id="5179a-284">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

<br>

---
---

#### <a name="displaymessageformitemid"></a><span data-ttu-id="5179a-285">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="5179a-285">displayMessageForm(itemId)</span></span>

<span data-ttu-id="5179a-286">既存のメッセージを表示します。</span><span class="sxs-lookup"><span data-stu-id="5179a-286">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="5179a-287">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="5179a-287">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="5179a-288">`displayMessageForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存のメッセージを開きます。</span><span class="sxs-lookup"><span data-stu-id="5179a-288">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="5179a-289">Outlook on the web では、このメソッドはフォームの本文が 32 KB 以下の文字数の場合にのみ指定のフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="5179a-289">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="5179a-290">指定のアイテム識別子が既存のメッセージを表していない場合は、クライアント コンピューターにはメッセージは表示されず、エラー メッセージも返されません。</span><span class="sxs-lookup"><span data-stu-id="5179a-290">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="5179a-p109">予定を表す `itemId` を含む `displayMessageForm` を使用しないでください。既存の予定を表示するには、`displayAppointmentForm` メソッドを使用します。新しい予定を作成するフォームを表示するには、`displayNewAppointmentForm` メソッドを使用します。</span><span class="sxs-lookup"><span data-stu-id="5179a-p109">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5179a-293">パラメーター</span><span class="sxs-lookup"><span data-stu-id="5179a-293">Parameters</span></span>

|<span data-ttu-id="5179a-294">名前</span><span class="sxs-lookup"><span data-stu-id="5179a-294">Name</span></span>| <span data-ttu-id="5179a-295">種類</span><span class="sxs-lookup"><span data-stu-id="5179a-295">Type</span></span>| <span data-ttu-id="5179a-296">説明</span><span class="sxs-lookup"><span data-stu-id="5179a-296">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="5179a-297">String</span><span class="sxs-lookup"><span data-stu-id="5179a-297">String</span></span>|<span data-ttu-id="5179a-298">既存のメッセージの Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="5179a-298">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5179a-299">要件</span><span class="sxs-lookup"><span data-stu-id="5179a-299">Requirements</span></span>

|<span data-ttu-id="5179a-300">要件</span><span class="sxs-lookup"><span data-stu-id="5179a-300">Requirement</span></span>| <span data-ttu-id="5179a-301">値</span><span class="sxs-lookup"><span data-stu-id="5179a-301">Value</span></span>|
|---|---|
|[<span data-ttu-id="5179a-302">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="5179a-302">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5179a-303">1.0</span><span class="sxs-lookup"><span data-stu-id="5179a-303">1.0</span></span>|
|[<span data-ttu-id="5179a-304">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="5179a-304">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5179a-305">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5179a-305">ReadItem</span></span>|
|[<span data-ttu-id="5179a-306">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="5179a-306">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5179a-307">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="5179a-307">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5179a-308">例</span><span class="sxs-lookup"><span data-stu-id="5179a-308">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

<br>

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="5179a-309">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="5179a-309">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="5179a-310">新しい予定を作成するためのフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="5179a-310">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="5179a-311">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="5179a-311">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="5179a-p110">`displayNewAppointmentForm` メソッドを使用すると、ユーザーが新しい予定または会議を作成できるフォームが開きます。パラメーターを指定すると、予定のフォーム フィールドにパラメーターの内容が自動的に設定されます。</span><span class="sxs-lookup"><span data-stu-id="5179a-p110">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="5179a-p111">Outlook on the web およびモバイル デバイスでは、このメソッドは常に出席者フィールドが含まれるフォームを表示します。入力引数として出席者を指定しないと、このメソッドは **[保存]** ボタンのあるフォームを表示します。出席者を指定した場合には、フォームにその出席者と **[送信]** ボタンが表示されます。</span><span class="sxs-lookup"><span data-stu-id="5179a-p111">In Outlook on the web and mobile devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="5179a-p112">Outlook リッチ クライアントと Outlook RT で、`requiredAttendees`、`optionalAttendees`、または `resources` パラメーターに出席者またはリソースを指定し、このメソッドを実行すると、**[送信]** ボタンがある会議フォームが表示されます。受信者を指定せずにこのメソッドを実行すると、**[保存して閉じる]** ボタンがある予定フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="5179a-p112">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="5179a-319">パラメーターのいずれかが指定のサイズ制限を超える場合、または不明なパラメーター名が指定されている場合は、例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="5179a-319">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5179a-320">パラメーター</span><span class="sxs-lookup"><span data-stu-id="5179a-320">Parameters</span></span>

|<span data-ttu-id="5179a-321">名前</span><span class="sxs-lookup"><span data-stu-id="5179a-321">Name</span></span>| <span data-ttu-id="5179a-322">種類</span><span class="sxs-lookup"><span data-stu-id="5179a-322">Type</span></span>| <span data-ttu-id="5179a-323">説明</span><span class="sxs-lookup"><span data-stu-id="5179a-323">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="5179a-324">Object</span><span class="sxs-lookup"><span data-stu-id="5179a-324">Object</span></span> | <span data-ttu-id="5179a-325">新しい予定を記述するパラメーターのディクショナリ。</span><span class="sxs-lookup"><span data-stu-id="5179a-325">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="5179a-326">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)&gt;</span><span class="sxs-lookup"><span data-stu-id="5179a-326">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)&gt;</span></span> | <span data-ttu-id="5179a-p113">予定に必要な各出席者について、メール アドレスを含む文字列の配列、または `EmailAddressDetails` オブジェクトを含む配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="5179a-p113">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="5179a-329">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)&gt;</span><span class="sxs-lookup"><span data-stu-id="5179a-329">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)&gt;</span></span> | <span data-ttu-id="5179a-p114">予定の各任意出席者について、メール アドレスを含む文字列の配列、または `EmailAddressDetails` オブジェクトを含む配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="5179a-p114">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="5179a-332">日付</span><span class="sxs-lookup"><span data-stu-id="5179a-332">Date</span></span> | <span data-ttu-id="5179a-333">予定の開始日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="5179a-333">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="5179a-334">日付</span><span class="sxs-lookup"><span data-stu-id="5179a-334">Date</span></span> | <span data-ttu-id="5179a-335">予定の終了日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="5179a-335">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="5179a-336">String</span><span class="sxs-lookup"><span data-stu-id="5179a-336">String</span></span> | <span data-ttu-id="5179a-p115">予定の場所を含む文字列。文字列は最大 255 文字に制限されます。</span><span class="sxs-lookup"><span data-stu-id="5179a-p115">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="5179a-339">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="5179a-339">Array.&lt;String&gt;</span></span> | <span data-ttu-id="5179a-p116">予定に必要なリソースを含む文字列の配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="5179a-p116">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="5179a-342">String</span><span class="sxs-lookup"><span data-stu-id="5179a-342">String</span></span> | <span data-ttu-id="5179a-p117">予定の件名を含む文字列です。文字列は最大 255 文字に制限されます。</span><span class="sxs-lookup"><span data-stu-id="5179a-p117">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="5179a-345">String</span><span class="sxs-lookup"><span data-stu-id="5179a-345">String</span></span> | <span data-ttu-id="5179a-p118">予定の本文。本文の内容は、最大サイズが 32 KB に制限されます。</span><span class="sxs-lookup"><span data-stu-id="5179a-p118">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="5179a-348">要件</span><span class="sxs-lookup"><span data-stu-id="5179a-348">Requirements</span></span>

|<span data-ttu-id="5179a-349">要件</span><span class="sxs-lookup"><span data-stu-id="5179a-349">Requirement</span></span>| <span data-ttu-id="5179a-350">値</span><span class="sxs-lookup"><span data-stu-id="5179a-350">Value</span></span>|
|---|---|
|[<span data-ttu-id="5179a-351">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="5179a-351">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5179a-352">1.0</span><span class="sxs-lookup"><span data-stu-id="5179a-352">1.0</span></span>|
|[<span data-ttu-id="5179a-353">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="5179a-353">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5179a-354">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5179a-354">ReadItem</span></span>|
|[<span data-ttu-id="5179a-355">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="5179a-355">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5179a-356">読み取り</span><span class="sxs-lookup"><span data-stu-id="5179a-356">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5179a-357">例</span><span class="sxs-lookup"><span data-stu-id="5179a-357">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="5179a-358">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="5179a-358">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="5179a-359">Exchange Server から添付ファイルやアイテムを取得するために使うトークンを含む文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="5179a-359">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="5179a-p119">`getCallbackTokenAsync` メソッドは、ユーザーのメールボックスをホストする Exchange Server から不透明なトークンを取得する非同期の呼び出しを行います。コールバック トークンの有効期間は 5 分です。</span><span class="sxs-lookup"><span data-stu-id="5179a-p119">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="5179a-362">トークンと、添付ファイル識別子またはアイテム識別子の両方をサードパーティのシステムに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="5179a-362">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="5179a-363">サードパーティシステムは、トークンをベアラー認証トークンとして使用して、Exchange Web サービス (EWS) の[Getattachment](/exchange/client-developer/web-service-reference/getattachment-operation)操作または[GetItem](/exchange/client-developer/web-service-reference/getitem-operation)操作を呼び出して、添付ファイルまたはアイテムを返します。</span><span class="sxs-lookup"><span data-stu-id="5179a-363">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item.</span></span> <span data-ttu-id="5179a-364">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="5179a-364">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="5179a-365">読み取りモード`getCallbackTokenAsync`でメソッドを呼び出すには、 **ReadItem**の最低限のアクセス許可レベルが必要です。</span><span class="sxs-lookup"><span data-stu-id="5179a-365">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="5179a-366">新規`getCallbackTokenAsync`作成モードで呼び出しを行うには、アイテムを保存しておく必要があります。</span><span class="sxs-lookup"><span data-stu-id="5179a-366">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="5179a-367">この[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback)メソッドには、 **readwriteitem**の最小アクセス許可レベルが必要です。</span><span class="sxs-lookup"><span data-stu-id="5179a-367">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5179a-368">パラメーター</span><span class="sxs-lookup"><span data-stu-id="5179a-368">Parameters</span></span>

|<span data-ttu-id="5179a-369">名前</span><span class="sxs-lookup"><span data-stu-id="5179a-369">Name</span></span>| <span data-ttu-id="5179a-370">種類</span><span class="sxs-lookup"><span data-stu-id="5179a-370">Type</span></span>| <span data-ttu-id="5179a-371">属性</span><span class="sxs-lookup"><span data-stu-id="5179a-371">Attributes</span></span>| <span data-ttu-id="5179a-372">説明</span><span class="sxs-lookup"><span data-stu-id="5179a-372">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="5179a-373">function</span><span class="sxs-lookup"><span data-stu-id="5179a-373">function</span></span>||<span data-ttu-id="5179a-374">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="5179a-374">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="5179a-375">トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="5179a-375">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="5179a-376">エラーが発生した場合、 `asyncResult.error` および `asyncResult.diagnostics` のプロパティで追加情報が提供される場合があります。</span><span class="sxs-lookup"><span data-stu-id="5179a-376">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="5179a-377">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="5179a-377">Object</span></span>| <span data-ttu-id="5179a-378">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="5179a-378">&lt;optional&gt;</span></span>|<span data-ttu-id="5179a-379">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="5179a-379">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="5179a-380">エラー</span><span class="sxs-lookup"><span data-stu-id="5179a-380">Errors</span></span>

|<span data-ttu-id="5179a-381">エラー コード</span><span class="sxs-lookup"><span data-stu-id="5179a-381">Error code</span></span>|<span data-ttu-id="5179a-382">説明</span><span class="sxs-lookup"><span data-stu-id="5179a-382">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="5179a-383">要求が失敗しました。</span><span class="sxs-lookup"><span data-stu-id="5179a-383">The request has failed.</span></span> <span data-ttu-id="5179a-384">HTTP エラーコードの diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="5179a-384">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="5179a-385">Exchange サーバーがエラーを返しました。</span><span class="sxs-lookup"><span data-stu-id="5179a-385">The Exchange server returned an error.</span></span> <span data-ttu-id="5179a-386">詳細については、diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="5179a-386">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="5179a-387">ユーザーはネットワークに接続されていません。</span><span class="sxs-lookup"><span data-stu-id="5179a-387">The user is no longer connected to the network.</span></span> <span data-ttu-id="5179a-388">ネットワーク接続を確認し、やり直してください。</span><span class="sxs-lookup"><span data-stu-id="5179a-388">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5179a-389">要件</span><span class="sxs-lookup"><span data-stu-id="5179a-389">Requirements</span></span>

|<span data-ttu-id="5179a-390">要件</span><span class="sxs-lookup"><span data-stu-id="5179a-390">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="5179a-391">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="5179a-391">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5179a-392">1.0</span><span class="sxs-lookup"><span data-stu-id="5179a-392">1.0</span></span> | <span data-ttu-id="5179a-393">1.3</span><span class="sxs-lookup"><span data-stu-id="5179a-393">1.3</span></span> |
|[<span data-ttu-id="5179a-394">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="5179a-394">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5179a-395">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5179a-395">ReadItem</span></span> | <span data-ttu-id="5179a-396">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5179a-396">ReadItem</span></span> |
|[<span data-ttu-id="5179a-397">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="5179a-397">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5179a-398">読み取り</span><span class="sxs-lookup"><span data-stu-id="5179a-398">Read</span></span> | <span data-ttu-id="5179a-399">作成</span><span class="sxs-lookup"><span data-stu-id="5179a-399">Compose</span></span> |

##### <a name="example"></a><span data-ttu-id="5179a-400">例</span><span class="sxs-lookup"><span data-stu-id="5179a-400">Example</span></span>

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

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="5179a-401">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="5179a-401">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="5179a-402">ユーザーと Office アドインを識別するトークンを取得します。</span><span class="sxs-lookup"><span data-stu-id="5179a-402">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="5179a-403">`getUserIdentityTokenAsync` メソッドは、[アドインとユーザーをサード パーティのシステムで識別して認証](/outlook/add-ins/authentication)することのできるトークンを返します。</span><span class="sxs-lookup"><span data-stu-id="5179a-403">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="5179a-404">パラメーター</span><span class="sxs-lookup"><span data-stu-id="5179a-404">Parameters</span></span>

|<span data-ttu-id="5179a-405">名前</span><span class="sxs-lookup"><span data-stu-id="5179a-405">Name</span></span>| <span data-ttu-id="5179a-406">種類</span><span class="sxs-lookup"><span data-stu-id="5179a-406">Type</span></span>| <span data-ttu-id="5179a-407">属性</span><span class="sxs-lookup"><span data-stu-id="5179a-407">Attributes</span></span>| <span data-ttu-id="5179a-408">説明</span><span class="sxs-lookup"><span data-stu-id="5179a-408">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="5179a-409">function</span><span class="sxs-lookup"><span data-stu-id="5179a-409">function</span></span>||<span data-ttu-id="5179a-410">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="5179a-410">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="5179a-411">トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="5179a-411">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="5179a-412">エラーが発生した場合、 `asyncResult.error` および `asyncResult.diagnostics` のプロパティで追加情報が提供される場合があります。</span><span class="sxs-lookup"><span data-stu-id="5179a-412">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="5179a-413">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="5179a-413">Object</span></span>| <span data-ttu-id="5179a-414">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="5179a-414">&lt;optional&gt;</span></span>|<span data-ttu-id="5179a-415">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="5179a-415">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="5179a-416">エラー</span><span class="sxs-lookup"><span data-stu-id="5179a-416">Errors</span></span>

|<span data-ttu-id="5179a-417">エラー コード</span><span class="sxs-lookup"><span data-stu-id="5179a-417">Error code</span></span>|<span data-ttu-id="5179a-418">説明</span><span class="sxs-lookup"><span data-stu-id="5179a-418">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="5179a-419">要求が失敗しました。</span><span class="sxs-lookup"><span data-stu-id="5179a-419">The request has failed.</span></span> <span data-ttu-id="5179a-420">HTTP エラーコードの diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="5179a-420">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="5179a-421">Exchange サーバーがエラーを返しました。</span><span class="sxs-lookup"><span data-stu-id="5179a-421">The Exchange server returned an error.</span></span> <span data-ttu-id="5179a-422">詳細については、diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="5179a-422">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="5179a-423">ユーザーはネットワークに接続されていません。</span><span class="sxs-lookup"><span data-stu-id="5179a-423">The user is no longer connected to the network.</span></span> <span data-ttu-id="5179a-424">ネットワーク接続を確認し、やり直してください。</span><span class="sxs-lookup"><span data-stu-id="5179a-424">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5179a-425">要件</span><span class="sxs-lookup"><span data-stu-id="5179a-425">Requirements</span></span>

|<span data-ttu-id="5179a-426">要件</span><span class="sxs-lookup"><span data-stu-id="5179a-426">Requirement</span></span>| <span data-ttu-id="5179a-427">値</span><span class="sxs-lookup"><span data-stu-id="5179a-427">Value</span></span>|
|---|---|
|[<span data-ttu-id="5179a-428">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="5179a-428">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5179a-429">1.0</span><span class="sxs-lookup"><span data-stu-id="5179a-429">1.0</span></span>|
|[<span data-ttu-id="5179a-430">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="5179a-430">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5179a-431">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5179a-431">ReadItem</span></span>|
|[<span data-ttu-id="5179a-432">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="5179a-432">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5179a-433">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="5179a-433">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5179a-434">例</span><span class="sxs-lookup"><span data-stu-id="5179a-434">Example</span></span>

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

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="5179a-435">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="5179a-435">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="5179a-436">ユーザーのメールボックスをホストしている Exchange サーバー上の Exchange Web サービス (EWS) のサービスに対して非同期の要求を行います。</span><span class="sxs-lookup"><span data-stu-id="5179a-436">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="5179a-437">このメソッドは、次のシナリオではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="5179a-437">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="5179a-438">Outlook on iOS または Android の場合</span><span class="sxs-lookup"><span data-stu-id="5179a-438">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="5179a-439">アドインが Gmail のメールボックスに読み込まれる場合</span><span class="sxs-lookup"><span data-stu-id="5179a-439">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="5179a-440">このような場合は、アドインでは [REST API を使用](/outlook/add-ins/use-rest-api)して、代わりにユーザーのメールボックスにアクセスする必要があります。</span><span class="sxs-lookup"><span data-stu-id="5179a-440">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="5179a-p128">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange. See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span><span class="sxs-lookup"><span data-stu-id="5179a-p128">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange. See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="5179a-443">`makeEwsRequestAsync` メソッドでは、フォルダー関連アイテムを要求できません。</span><span class="sxs-lookup"><span data-stu-id="5179a-443">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="5179a-444">XML 要求では UTF-8 エンコードを指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="5179a-444">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="5179a-p129">`makeEwsRequestAsync` メソッドを使用するには、アドインに **ReadWriteMailbox** アクセス許可が必要です。**ReadWriteMailbox** アクセス許可と、`makeEwsRequestAsync` メソッドで呼び出せる EWS 操作の使い方については、「[ユーザーのメールボックスへのメール アドイン アクセスのアクセス許可を指定する](/outlook/add-ins/understanding-outlook-add-in-permissions)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="5179a-p129">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="5179a-447">サーバー管理者は、クライアント アクセス サーバーの EWS ディレクトリで `OAuthAuthentication` を true に設定して、`makeEwsRequestAsync` メソッドで EWS 要求を行うことができるようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="5179a-447">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="5179a-448">バージョンの相違点</span><span class="sxs-lookup"><span data-stu-id="5179a-448">Version differences</span></span>

<span data-ttu-id="5179a-449">バージョン 15.0.4535.1004 より前のバージョンの Outlook で実行しているメール アプリで `makeEwsRequestAsync` メソッドを使う場合は、エンコード値を `ISO-8859-1` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="5179a-449">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="5179a-p130">Outlook on the web でメール アプリを実行している場合は、エンコード値を設定する必要はありません。mailbox.diagnostics.hostName プロパティを使って、メール アプリを Outlook で実行しているのか、Outlook on the web で実行しているのかを確認できます。mailbox.diagnostics.hostVersion プロパティを使って、どのバージョンの Outlook を使って実行しているかを確認できます。</span><span class="sxs-lookup"><span data-stu-id="5179a-p130">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5179a-453">パラメーター</span><span class="sxs-lookup"><span data-stu-id="5179a-453">Parameters</span></span>

|<span data-ttu-id="5179a-454">名前</span><span class="sxs-lookup"><span data-stu-id="5179a-454">Name</span></span>| <span data-ttu-id="5179a-455">種類</span><span class="sxs-lookup"><span data-stu-id="5179a-455">Type</span></span>| <span data-ttu-id="5179a-456">属性</span><span class="sxs-lookup"><span data-stu-id="5179a-456">Attributes</span></span>| <span data-ttu-id="5179a-457">説明</span><span class="sxs-lookup"><span data-stu-id="5179a-457">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="5179a-458">String</span><span class="sxs-lookup"><span data-stu-id="5179a-458">String</span></span>||<span data-ttu-id="5179a-459">EWS 要求です。</span><span class="sxs-lookup"><span data-stu-id="5179a-459">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="5179a-460">function</span><span class="sxs-lookup"><span data-stu-id="5179a-460">function</span></span>||<span data-ttu-id="5179a-461">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="5179a-461">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="5179a-p131">The XML result of the EWS call is provided as a string in the `asyncResult.value` property. If the result exceeds 1 MB in size, an error message is returned instead.</span><span class="sxs-lookup"><span data-stu-id="5179a-p131">The XML result of the EWS call is provided as a string in the `asyncResult.value` property. If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="5179a-464">Object</span><span class="sxs-lookup"><span data-stu-id="5179a-464">Object</span></span>| <span data-ttu-id="5179a-465">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="5179a-465">&lt;optional&gt;</span></span>|<span data-ttu-id="5179a-466">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="5179a-466">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5179a-467">要件</span><span class="sxs-lookup"><span data-stu-id="5179a-467">Requirements</span></span>

|<span data-ttu-id="5179a-468">要件</span><span class="sxs-lookup"><span data-stu-id="5179a-468">Requirement</span></span>| <span data-ttu-id="5179a-469">値</span><span class="sxs-lookup"><span data-stu-id="5179a-469">Value</span></span>|
|---|---|
|[<span data-ttu-id="5179a-470">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="5179a-470">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5179a-471">1.0</span><span class="sxs-lookup"><span data-stu-id="5179a-471">1.0</span></span>|
|[<span data-ttu-id="5179a-472">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="5179a-472">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5179a-473">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="5179a-473">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="5179a-474">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="5179a-474">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5179a-475">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="5179a-475">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5179a-476">例</span><span class="sxs-lookup"><span data-stu-id="5179a-476">Example</span></span>

<span data-ttu-id="5179a-477">次の例は、`makeEwsRequestAsync` を呼び出し、`GetItem` 操作を使って項目の件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="5179a-477">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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
