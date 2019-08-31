---
title: Office. メールボックス要件セット1.3
description: ''
ms.date: 08/30/2019
localization_priority: Normal
ms.openlocfilehash: 05b7d82e036cc29526c18bf97c6a1472778c1959
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696233"
---
# <a name="mailbox"></a><span data-ttu-id="58095-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="58095-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="58095-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="58095-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="58095-104">Microsoft Outlook の Outlook アドインオブジェクトモデルへのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="58095-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="58095-105">要件</span><span class="sxs-lookup"><span data-stu-id="58095-105">Requirements</span></span>

|<span data-ttu-id="58095-106">要件</span><span class="sxs-lookup"><span data-stu-id="58095-106">Requirement</span></span>| <span data-ttu-id="58095-107">値</span><span class="sxs-lookup"><span data-stu-id="58095-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="58095-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="58095-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="58095-109">1.0</span><span class="sxs-lookup"><span data-stu-id="58095-109">1.0</span></span>|
|[<span data-ttu-id="58095-110">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="58095-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="58095-111">制限あり</span><span class="sxs-lookup"><span data-stu-id="58095-111">Restricted</span></span>|
|[<span data-ttu-id="58095-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="58095-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="58095-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="58095-113">Compose or Read</span></span>|

<span data-ttu-id="58095-114">| [Ewsurl](#ewsurl-string) |メンバ | |[Converttoewsid](#converttoewsiditemid-restversion--string) |メソッド | |[Converttolocalclienttime](#converttolocalclienttimetimevalue--localclienttime) |メソッド | |[Converttorestid](#converttorestiditemid-restversion--string) |メソッド | |[convertToUtcClientTime](#converttoutcclienttimeinput--date) |メソッド | |[displayAppointmentForm](#displayappointmentformitemid) |メソッド | |[Displaymessageform](#displaymessageformitemid) |メソッド | |[displayNewAppointmentForm](#displaynewappointmentformparameters) |メソッド | |[Get/Tokenasync](#getcallbacktokenasynccallback-usercontext) |メソッド | |[Getuseridentity Tokenasync](#getuseridentitytokenasynccallback-usercontext) |メソッド | |[makeEwsRequestAsync](#makeewsrequestasyncdata-callback-usercontext) |メソッド |</span><span class="sxs-lookup"><span data-stu-id="58095-114">| [ewsUrl](#ewsurl-string) | Member | | [convertToEwsId](#converttoewsiditemid-restversion--string) | Method | | [convertToLocalClientTime](#converttolocalclienttimetimevalue--localclienttime) | Method | | [convertToRestId](#converttorestiditemid-restversion--string) | Method | | [convertToUtcClientTime](#converttoutcclienttimeinput--date) | Method | | [displayAppointmentForm](#displayappointmentformitemid) | Method | | [displayMessageForm](#displaymessageformitemid) | Method | | [displayNewAppointmentForm](#displaynewappointmentformparameters) | Method | | [getCallbackTokenAsync](#getcallbacktokenasynccallback-usercontext) | Method | | [getUserIdentityTokenAsync](#getuseridentitytokenasynccallback-usercontext) | Method | | [makeEwsRequestAsync](#makeewsrequestasyncdata-callback-usercontext) | Method |</span></span>

### <a name="namespaces"></a><span data-ttu-id="58095-115">名前空間</span><span class="sxs-lookup"><span data-stu-id="58095-115">Namespaces</span></span>

<span data-ttu-id="58095-116">[diagnostics](Office.context.mailbox.diagnostics.md):Outlook アドインに診断情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="58095-116">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="58095-117">[item](Office.context.mailbox.item.md):Outlook アドインのメッセージや予定にアクセスするためのメソッドとプロパティを提供します。</span><span class="sxs-lookup"><span data-stu-id="58095-117">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="58095-118">[userProfile](Office.context.mailbox.userProfile.md):Outlook アドインのユーザーに関する情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="58095-118">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="58095-119">メンバー</span><span class="sxs-lookup"><span data-stu-id="58095-119">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="58095-120">ewsUrl: String</span><span class="sxs-lookup"><span data-stu-id="58095-120">ewsUrl: String</span></span>

<span data-ttu-id="58095-121">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account.</span><span class="sxs-lookup"><span data-stu-id="58095-121">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account.</span></span> <span data-ttu-id="58095-122">Read mode only.</span><span class="sxs-lookup"><span data-stu-id="58095-122">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="58095-123">このメンバーは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="58095-123">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="58095-p102">
  `ewsUrl\` 値は、リモート サービスで、ユーザーのメールボックスに EWS 呼び出しを行うために使うことができます。たとえば、[選択したアイテムから添付ファイルを取得する](/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="58095-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="58095-126">アプリが閲覧モードで `ewsUrl` メンバーを呼び出すには、アプリのマニフェスト内に **ReadItem** アクセス許可が指定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="58095-126">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="58095-p103">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出してから、`ewsUrl` メンバーを使用する必要があります。アプリには、`saveAsync` メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="58095-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="58095-129">型</span><span class="sxs-lookup"><span data-stu-id="58095-129">Type</span></span>

*   <span data-ttu-id="58095-130">String</span><span class="sxs-lookup"><span data-stu-id="58095-130">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="58095-131">要件</span><span class="sxs-lookup"><span data-stu-id="58095-131">Requirements</span></span>

|<span data-ttu-id="58095-132">要件</span><span class="sxs-lookup"><span data-stu-id="58095-132">Requirement</span></span>| <span data-ttu-id="58095-133">値</span><span class="sxs-lookup"><span data-stu-id="58095-133">Value</span></span>|
|---|---|
|[<span data-ttu-id="58095-134">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="58095-134">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="58095-135">1.0</span><span class="sxs-lookup"><span data-stu-id="58095-135">1.0</span></span>|
|[<span data-ttu-id="58095-136">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="58095-136">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="58095-137">ReadItem</span><span class="sxs-lookup"><span data-stu-id="58095-137">ReadItem</span></span>|
|[<span data-ttu-id="58095-138">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="58095-138">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="58095-139">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="58095-139">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="58095-140">メソッド</span><span class="sxs-lookup"><span data-stu-id="58095-140">Methods</span></span>

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="58095-141">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="58095-141">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="58095-142">REST 形式のアイテム ID を EWS 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="58095-142">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="58095-143">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="58095-143">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="58095-p104">REST API ([Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) や [Microsoft Graph](https://graph.microsoft.io/) など) で取得されたアイテム ID は、Exchange Web サービス (EWS) に使用される形式とは異なる形式を使用します。`convertToEwsId` メソッドは、REST 形式の ID を EWS 用の適切な形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="58095-p104">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="58095-146">パラメーター</span><span class="sxs-lookup"><span data-stu-id="58095-146">Parameters</span></span>

|<span data-ttu-id="58095-147">名前</span><span class="sxs-lookup"><span data-stu-id="58095-147">Name</span></span>| <span data-ttu-id="58095-148">種類</span><span class="sxs-lookup"><span data-stu-id="58095-148">Type</span></span>| <span data-ttu-id="58095-149">説明</span><span class="sxs-lookup"><span data-stu-id="58095-149">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="58095-150">String</span><span class="sxs-lookup"><span data-stu-id="58095-150">String</span></span>|<span data-ttu-id="58095-151">Outlook REST API 形式のアイテム ID</span><span class="sxs-lookup"><span data-stu-id="58095-151">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="58095-152">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="58095-152">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.3)|<span data-ttu-id="58095-153">アイテム ID の取得に使用された Outlook REST API のバージョンを示す値。</span><span class="sxs-lookup"><span data-stu-id="58095-153">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="58095-154">要件</span><span class="sxs-lookup"><span data-stu-id="58095-154">Requirements</span></span>

|<span data-ttu-id="58095-155">要件</span><span class="sxs-lookup"><span data-stu-id="58095-155">Requirement</span></span>| <span data-ttu-id="58095-156">値</span><span class="sxs-lookup"><span data-stu-id="58095-156">Value</span></span>|
|---|---|
|[<span data-ttu-id="58095-157">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="58095-157">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="58095-158">1.3</span><span class="sxs-lookup"><span data-stu-id="58095-158">1.3</span></span>|
|[<span data-ttu-id="58095-159">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="58095-159">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="58095-160">制限あり</span><span class="sxs-lookup"><span data-stu-id="58095-160">Restricted</span></span>|
|[<span data-ttu-id="58095-161">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="58095-161">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="58095-162">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="58095-162">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="58095-163">戻り値:</span><span class="sxs-lookup"><span data-stu-id="58095-163">Returns:</span></span>

<span data-ttu-id="58095-164">型:String</span><span class="sxs-lookup"><span data-stu-id="58095-164">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="58095-165">例</span><span class="sxs-lookup"><span data-stu-id="58095-165">Example</span></span>

```js
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-13"></a><span data-ttu-id="58095-166">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.3)}</span><span class="sxs-lookup"><span data-stu-id="58095-166">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.3)}</span></span>

<span data-ttu-id="58095-167">クライアントのローカル時間で時間情報が含まれている辞書を取得します。</span><span class="sxs-lookup"><span data-stu-id="58095-167">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="58095-168">デスクトップまたは web 上の Outlook 用メールアプリは、日付と時刻に異なるタイムゾーンを使用できます。</span><span class="sxs-lookup"><span data-stu-id="58095-168">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times.</span></span> <span data-ttu-id="58095-169">デスクトップ上の Outlook では、クライアントコンピューターのタイムゾーンが使用されます。Outlook on the web では、Exchange 管理センター (EAC) で設定されているタイムゾーンが使用されます。</span><span class="sxs-lookup"><span data-stu-id="58095-169">Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC).</span></span> <span data-ttu-id="58095-170">日付と時刻の値を処理して、ユーザーインターフェイスに表示される値が、ユーザーが期待するタイムゾーンに常に一致するようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="58095-170">You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="58095-171">デスクトップクライアント上の Outlook でメールアプリが実行されている`convertToLocalClientTime`場合、このメソッドは、クライアントコンピューターのタイムゾーンに設定された値を持つ dictionary オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="58095-171">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone.</span></span> <span data-ttu-id="58095-172">メールアプリが web 上の Outlook で実行されている`convertToLocalClientTime`場合、このメソッドは、EAC で指定されたタイムゾーンに設定された値を持つ dictionary オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="58095-172">If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="58095-173">パラメーター</span><span class="sxs-lookup"><span data-stu-id="58095-173">Parameters</span></span>

|<span data-ttu-id="58095-174">名前</span><span class="sxs-lookup"><span data-stu-id="58095-174">Name</span></span>| <span data-ttu-id="58095-175">種類</span><span class="sxs-lookup"><span data-stu-id="58095-175">Type</span></span>| <span data-ttu-id="58095-176">説明</span><span class="sxs-lookup"><span data-stu-id="58095-176">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="58095-177">日付</span><span class="sxs-lookup"><span data-stu-id="58095-177">Date</span></span>|<span data-ttu-id="58095-178">日付オブジェクト</span><span class="sxs-lookup"><span data-stu-id="58095-178">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="58095-179">要件</span><span class="sxs-lookup"><span data-stu-id="58095-179">Requirements</span></span>

|<span data-ttu-id="58095-180">要件</span><span class="sxs-lookup"><span data-stu-id="58095-180">Requirement</span></span>| <span data-ttu-id="58095-181">値</span><span class="sxs-lookup"><span data-stu-id="58095-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="58095-182">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="58095-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="58095-183">1.0</span><span class="sxs-lookup"><span data-stu-id="58095-183">1.0</span></span>|
|[<span data-ttu-id="58095-184">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="58095-184">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="58095-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="58095-185">ReadItem</span></span>|
|[<span data-ttu-id="58095-186">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="58095-186">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="58095-187">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="58095-187">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="58095-188">戻り値:</span><span class="sxs-lookup"><span data-stu-id="58095-188">Returns:</span></span>

<span data-ttu-id="58095-189">型:[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="58095-189">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.3)</span></span>

<br>

---
---

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="58095-190">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="58095-190">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="58095-191">EWS 形式のアイテム ID を REST 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="58095-191">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="58095-192">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="58095-192">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="58095-p107">EWS または `itemId` プロパティで取得されるアイテム ID は、REST API ([Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) や [Microsoft Graph](https://graph.microsoft.io/) など) に使用される形式とは異なる形式を使用します。`convertToRestId` メソッドは、EWS 形式の ID を REST 用の適切な形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="58095-p107">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="58095-195">パラメーター</span><span class="sxs-lookup"><span data-stu-id="58095-195">Parameters</span></span>

|<span data-ttu-id="58095-196">名前</span><span class="sxs-lookup"><span data-stu-id="58095-196">Name</span></span>| <span data-ttu-id="58095-197">種類</span><span class="sxs-lookup"><span data-stu-id="58095-197">Type</span></span>| <span data-ttu-id="58095-198">説明</span><span class="sxs-lookup"><span data-stu-id="58095-198">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="58095-199">String</span><span class="sxs-lookup"><span data-stu-id="58095-199">String</span></span>|<span data-ttu-id="58095-200">Exchange Web サービス (EWS) 形式のアイテム ID</span><span class="sxs-lookup"><span data-stu-id="58095-200">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="58095-201">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="58095-201">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.3)|<span data-ttu-id="58095-202">変換後の ID を使用する Outlook REST API のバージョンを示す値。</span><span class="sxs-lookup"><span data-stu-id="58095-202">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="58095-203">要件</span><span class="sxs-lookup"><span data-stu-id="58095-203">Requirements</span></span>

|<span data-ttu-id="58095-204">要件</span><span class="sxs-lookup"><span data-stu-id="58095-204">Requirement</span></span>| <span data-ttu-id="58095-205">値</span><span class="sxs-lookup"><span data-stu-id="58095-205">Value</span></span>|
|---|---|
|[<span data-ttu-id="58095-206">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="58095-206">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="58095-207">1.3</span><span class="sxs-lookup"><span data-stu-id="58095-207">1.3</span></span>|
|[<span data-ttu-id="58095-208">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="58095-208">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="58095-209">制限あり</span><span class="sxs-lookup"><span data-stu-id="58095-209">Restricted</span></span>|
|[<span data-ttu-id="58095-210">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="58095-210">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="58095-211">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="58095-211">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="58095-212">戻り値:</span><span class="sxs-lookup"><span data-stu-id="58095-212">Returns:</span></span>

<span data-ttu-id="58095-213">型:String</span><span class="sxs-lookup"><span data-stu-id="58095-213">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="58095-214">例</span><span class="sxs-lookup"><span data-stu-id="58095-214">Example</span></span>

```js
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="58095-215">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="58095-215">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="58095-216">時間情報が含まれているディクショナリから日付オブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="58095-216">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="58095-217">`convertToUtcClientTime` メソッドは、ローカルの日付と時刻を含むディクショナリを、ローカルの日付と時刻の正しい値を持つ日付オブジェクトに変換します。</span><span class="sxs-lookup"><span data-stu-id="58095-217">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="58095-218">パラメーター</span><span class="sxs-lookup"><span data-stu-id="58095-218">Parameters</span></span>

|<span data-ttu-id="58095-219">名前</span><span class="sxs-lookup"><span data-stu-id="58095-219">Name</span></span>| <span data-ttu-id="58095-220">型</span><span class="sxs-lookup"><span data-stu-id="58095-220">Type</span></span>| <span data-ttu-id="58095-221">説明</span><span class="sxs-lookup"><span data-stu-id="58095-221">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="58095-222">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="58095-222">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.3)|<span data-ttu-id="58095-223">変換するローカル時刻の値。</span><span class="sxs-lookup"><span data-stu-id="58095-223">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="58095-224">要件</span><span class="sxs-lookup"><span data-stu-id="58095-224">Requirements</span></span>

|<span data-ttu-id="58095-225">要件</span><span class="sxs-lookup"><span data-stu-id="58095-225">Requirement</span></span>| <span data-ttu-id="58095-226">値</span><span class="sxs-lookup"><span data-stu-id="58095-226">Value</span></span>|
|---|---|
|[<span data-ttu-id="58095-227">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="58095-227">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="58095-228">1.0</span><span class="sxs-lookup"><span data-stu-id="58095-228">1.0</span></span>|
|[<span data-ttu-id="58095-229">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="58095-229">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="58095-230">ReadItem</span><span class="sxs-lookup"><span data-stu-id="58095-230">ReadItem</span></span>|
|[<span data-ttu-id="58095-231">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="58095-231">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="58095-232">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="58095-232">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="58095-233">戻り値:</span><span class="sxs-lookup"><span data-stu-id="58095-233">Returns:</span></span>

<span data-ttu-id="58095-234">時間が UTC で表現された日付オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="58095-234">A Date object with the time expressed in UTC.</span></span>

<span data-ttu-id="58095-235">型: Date</span><span class="sxs-lookup"><span data-stu-id="58095-235">Type: Date</span></span>

##### <a name="example"></a><span data-ttu-id="58095-236">例</span><span class="sxs-lookup"><span data-stu-id="58095-236">Example</span></span>

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

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="58095-237">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="58095-237">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="58095-238">既存の予定を表示します。</span><span class="sxs-lookup"><span data-stu-id="58095-238">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="58095-239">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="58095-239">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="58095-240">`displayAppointmentForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存の予定を開きます。</span><span class="sxs-lookup"><span data-stu-id="58095-240">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="58095-241">Outlook on the Mac では、このメソッドを使用して、定期的なアイテムの一部ではない単一の予定を表示したり、定期的なアイテムのマスター予定を表示したりすることはできませんが、一連のインスタンスを表示することはできません。</span><span class="sxs-lookup"><span data-stu-id="58095-241">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series.</span></span> <span data-ttu-id="58095-242">これは、Mac 上の Outlook では、定期的なアイテムのインスタンスのプロパティ (アイテム ID を含む) にアクセスできないためです。</span><span class="sxs-lookup"><span data-stu-id="58095-242">This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="58095-243">Web 上の Outlook では、フォームの本文が 32 KB 以下の文字である場合にのみ、このメソッドは指定されたフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="58095-243">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="58095-244">指定のアイテム識別子が既存の予定を表していない場合は、クライアント コンピューターまたはデバイスで空のウィンドウが開き、エラー メッセージは返されません。</span><span class="sxs-lookup"><span data-stu-id="58095-244">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="58095-245">パラメーター</span><span class="sxs-lookup"><span data-stu-id="58095-245">Parameters</span></span>

|<span data-ttu-id="58095-246">名前</span><span class="sxs-lookup"><span data-stu-id="58095-246">Name</span></span>| <span data-ttu-id="58095-247">型</span><span class="sxs-lookup"><span data-stu-id="58095-247">Type</span></span>| <span data-ttu-id="58095-248">説明</span><span class="sxs-lookup"><span data-stu-id="58095-248">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="58095-249">String</span><span class="sxs-lookup"><span data-stu-id="58095-249">String</span></span>|<span data-ttu-id="58095-250">既存の予定の Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="58095-250">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="58095-251">要件</span><span class="sxs-lookup"><span data-stu-id="58095-251">Requirements</span></span>

|<span data-ttu-id="58095-252">要件</span><span class="sxs-lookup"><span data-stu-id="58095-252">Requirement</span></span>| <span data-ttu-id="58095-253">値</span><span class="sxs-lookup"><span data-stu-id="58095-253">Value</span></span>|
|---|---|
|[<span data-ttu-id="58095-254">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="58095-254">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="58095-255">1.0</span><span class="sxs-lookup"><span data-stu-id="58095-255">1.0</span></span>|
|[<span data-ttu-id="58095-256">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="58095-256">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="58095-257">ReadItem</span><span class="sxs-lookup"><span data-stu-id="58095-257">ReadItem</span></span>|
|[<span data-ttu-id="58095-258">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="58095-258">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="58095-259">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="58095-259">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="58095-260">例</span><span class="sxs-lookup"><span data-stu-id="58095-260">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

<br>

---
---

#### <a name="displaymessageformitemid"></a><span data-ttu-id="58095-261">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="58095-261">displayMessageForm(itemId)</span></span>

<span data-ttu-id="58095-262">既存のメッセージを表示します。</span><span class="sxs-lookup"><span data-stu-id="58095-262">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="58095-263">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="58095-263">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="58095-264">`displayMessageForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存のメッセージを開きます。</span><span class="sxs-lookup"><span data-stu-id="58095-264">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="58095-265">Web 上の Outlook では、フォームの本文が 32 KB の文字数以下の場合にのみ、このメソッドは指定されたフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="58095-265">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="58095-266">指定のアイテム識別子が既存のメッセージを表していない場合は、クライアント コンピューターにはメッセージは表示されず、エラー メッセージも返されません。</span><span class="sxs-lookup"><span data-stu-id="58095-266">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="58095-p109">予定を表す `itemId` を含む `displayMessageForm` を使用しないでください。既存の予定を表示するには、`displayAppointmentForm` メソッドを使用します。新しい予定を作成するフォームを表示するには、`displayNewAppointmentForm` メソッドを使用します。</span><span class="sxs-lookup"><span data-stu-id="58095-p109">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="58095-269">パラメーター</span><span class="sxs-lookup"><span data-stu-id="58095-269">Parameters</span></span>

|<span data-ttu-id="58095-270">名前</span><span class="sxs-lookup"><span data-stu-id="58095-270">Name</span></span>| <span data-ttu-id="58095-271">型</span><span class="sxs-lookup"><span data-stu-id="58095-271">Type</span></span>| <span data-ttu-id="58095-272">説明</span><span class="sxs-lookup"><span data-stu-id="58095-272">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="58095-273">String</span><span class="sxs-lookup"><span data-stu-id="58095-273">String</span></span>|<span data-ttu-id="58095-274">既存のメッセージの Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="58095-274">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="58095-275">要件</span><span class="sxs-lookup"><span data-stu-id="58095-275">Requirements</span></span>

|<span data-ttu-id="58095-276">要件</span><span class="sxs-lookup"><span data-stu-id="58095-276">Requirement</span></span>| <span data-ttu-id="58095-277">値</span><span class="sxs-lookup"><span data-stu-id="58095-277">Value</span></span>|
|---|---|
|[<span data-ttu-id="58095-278">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="58095-278">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="58095-279">1.0</span><span class="sxs-lookup"><span data-stu-id="58095-279">1.0</span></span>|
|[<span data-ttu-id="58095-280">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="58095-280">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="58095-281">ReadItem</span><span class="sxs-lookup"><span data-stu-id="58095-281">ReadItem</span></span>|
|[<span data-ttu-id="58095-282">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="58095-282">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="58095-283">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="58095-283">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="58095-284">例</span><span class="sxs-lookup"><span data-stu-id="58095-284">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

<br>

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="58095-285">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="58095-285">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="58095-286">新しい予定を作成するためのフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="58095-286">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="58095-287">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="58095-287">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="58095-p110">`displayNewAppointmentForm` メソッドを使用すると、ユーザーが新しい予定または会議を作成できるフォームが開きます。パラメーターを指定すると、予定のフォーム フィールドにパラメーターの内容が自動的に設定されます。</span><span class="sxs-lookup"><span data-stu-id="58095-p110">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="58095-290">Outlook on the web およびモバイルデバイスでは、このメソッドは常に出席者フィールドを含むフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="58095-290">In Outlook on the web and mobile devices, this method always displays a form with an attendees field.</span></span> <span data-ttu-id="58095-291">入力引数として出席者を指定しないと、このメソッドにより **[保存]** ボタンのあるフォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="58095-291">If you do not specify any attendees as input arguments, the method displays a form with a **Save** button.</span></span> <span data-ttu-id="58095-292">出席者を指定した場合には、フォームにその出席者と **[送信]** ボタンが表示されます。</span><span class="sxs-lookup"><span data-stu-id="58095-292">If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="58095-p112">Outlook リッチ クライアントと Outlook RT で、`requiredAttendees`、`optionalAttendees`、または `resources` パラメーターに出席者またはリソースを指定し、このメソッドを実行すると、**[送信]** ボタンがある会議フォームが表示されます。受信者を指定せずにこのメソッドを実行すると、**[保存して閉じる]** ボタンがある予定フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="58095-p112">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="58095-295">パラメーターのいずれかが指定のサイズ制限を超える場合、または不明なパラメーター名が指定されている場合は、例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="58095-295">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="58095-296">パラメーター</span><span class="sxs-lookup"><span data-stu-id="58095-296">Parameters</span></span>

|<span data-ttu-id="58095-297">名前</span><span class="sxs-lookup"><span data-stu-id="58095-297">Name</span></span>| <span data-ttu-id="58095-298">型</span><span class="sxs-lookup"><span data-stu-id="58095-298">Type</span></span>| <span data-ttu-id="58095-299">説明</span><span class="sxs-lookup"><span data-stu-id="58095-299">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="58095-300">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="58095-300">Object</span></span> | <span data-ttu-id="58095-301">新しい予定を記述するパラメーターのディクショナリ。</span><span class="sxs-lookup"><span data-stu-id="58095-301">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="58095-302">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)&gt;</span><span class="sxs-lookup"><span data-stu-id="58095-302">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)&gt;</span></span> | <span data-ttu-id="58095-p113">予定に必要な各出席者について、メール アドレスを含む文字列の配列、または `EmailAddressDetails` オブジェクトを含む配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="58095-p113">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="58095-305">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)&gt;</span><span class="sxs-lookup"><span data-stu-id="58095-305">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)&gt;</span></span> | <span data-ttu-id="58095-p114">予定の各任意出席者について、メール アドレスを含む文字列の配列、または `EmailAddressDetails` オブジェクトを含む配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="58095-p114">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="58095-308">日付</span><span class="sxs-lookup"><span data-stu-id="58095-308">Date</span></span> | <span data-ttu-id="58095-309">予定の開始日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="58095-309">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="58095-310">日付</span><span class="sxs-lookup"><span data-stu-id="58095-310">Date</span></span> | <span data-ttu-id="58095-311">予定の終了日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="58095-311">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="58095-312">String</span><span class="sxs-lookup"><span data-stu-id="58095-312">String</span></span> | <span data-ttu-id="58095-p115">予定の場所を含む文字列。文字列は最大 255 文字に制限されます。</span><span class="sxs-lookup"><span data-stu-id="58095-p115">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="58095-315">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="58095-315">Array.&lt;String&gt;</span></span> | <span data-ttu-id="58095-p116">予定に必要なリソースを含む文字列の配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="58095-p116">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="58095-318">String</span><span class="sxs-lookup"><span data-stu-id="58095-318">String</span></span> | <span data-ttu-id="58095-p117">予定の件名を含む文字列です。文字列は最大 255 文字に制限されます。</span><span class="sxs-lookup"><span data-stu-id="58095-p117">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="58095-321">String</span><span class="sxs-lookup"><span data-stu-id="58095-321">String</span></span> | <span data-ttu-id="58095-p118">予定の本文。本文の内容は、最大サイズが 32 KB に制限されます。</span><span class="sxs-lookup"><span data-stu-id="58095-p118">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="58095-324">要件</span><span class="sxs-lookup"><span data-stu-id="58095-324">Requirements</span></span>

|<span data-ttu-id="58095-325">要件</span><span class="sxs-lookup"><span data-stu-id="58095-325">Requirement</span></span>| <span data-ttu-id="58095-326">値</span><span class="sxs-lookup"><span data-stu-id="58095-326">Value</span></span>|
|---|---|
|[<span data-ttu-id="58095-327">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="58095-327">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="58095-328">1.0</span><span class="sxs-lookup"><span data-stu-id="58095-328">1.0</span></span>|
|[<span data-ttu-id="58095-329">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="58095-329">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="58095-330">ReadItem</span><span class="sxs-lookup"><span data-stu-id="58095-330">ReadItem</span></span>|
|[<span data-ttu-id="58095-331">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="58095-331">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="58095-332">読み取り</span><span class="sxs-lookup"><span data-stu-id="58095-332">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="58095-333">例</span><span class="sxs-lookup"><span data-stu-id="58095-333">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="58095-334">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="58095-334">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="58095-335">Exchange Server から添付ファイルやアイテムを取得するために使うトークンを含む文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="58095-335">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="58095-p119">`getCallbackTokenAsync` メソッドは、ユーザーのメールボックスをホストする Exchange Server から不透明なトークンを取得する非同期の呼び出しを行います。コールバック トークンの有効期間は 5 分です。</span><span class="sxs-lookup"><span data-stu-id="58095-p119">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="58095-p120">トークンと予定の識別子またはアイテムの識別子をサードパーティ システムに渡すことができます。サードパーティ システムは、トークンをベアラー承認トークンとして使用し、Exchange Web サービス (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) または [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) 操作を呼び出して、添付ファイルまたはアイテムを返します。たとえば、[選択したアイテムから添付ファイルを取得する](/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="58095-p120">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="58095-341">アプリが閲覧モードで `getCallbackTokenAsync` メソッドを呼び出すには、アプリのマニフェスト内に **ReadItem** アクセス許可が指定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="58095-341">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="58095-p121">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出してアイテムの識別子を `getCallbackTokenAsync` メソッドに渡す必要があります。アプリには、`saveAsync` メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="58095-p121">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="58095-344">パラメーター</span><span class="sxs-lookup"><span data-stu-id="58095-344">Parameters</span></span>

|<span data-ttu-id="58095-345">名前</span><span class="sxs-lookup"><span data-stu-id="58095-345">Name</span></span>| <span data-ttu-id="58095-346">型</span><span class="sxs-lookup"><span data-stu-id="58095-346">Type</span></span>| <span data-ttu-id="58095-347">属性</span><span class="sxs-lookup"><span data-stu-id="58095-347">Attributes</span></span>| <span data-ttu-id="58095-348">説明</span><span class="sxs-lookup"><span data-stu-id="58095-348">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="58095-349">function</span><span class="sxs-lookup"><span data-stu-id="58095-349">function</span></span>||<span data-ttu-id="58095-350">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="58095-350">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="58095-351">トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="58095-351">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="58095-352">エラーが発生した場合は`asyncResult.error` 、 `asyncResult.diagnostics`プロパティとプロパティによって追加情報が提供されることがあります。</span><span class="sxs-lookup"><span data-stu-id="58095-352">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="58095-353">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="58095-353">Object</span></span>| <span data-ttu-id="58095-354">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="58095-354">&lt;optional&gt;</span></span>|<span data-ttu-id="58095-355">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="58095-355">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="58095-356">エラー</span><span class="sxs-lookup"><span data-stu-id="58095-356">Errors</span></span>

|<span data-ttu-id="58095-357">エラー コード</span><span class="sxs-lookup"><span data-stu-id="58095-357">Error code</span></span>|<span data-ttu-id="58095-358">説明</span><span class="sxs-lookup"><span data-stu-id="58095-358">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="58095-359">要求が失敗しました。</span><span class="sxs-lookup"><span data-stu-id="58095-359">The request has failed.</span></span> <span data-ttu-id="58095-360">HTTP エラーコードについては、diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="58095-360">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="58095-361">Exchange サーバーがエラーを返しました。</span><span class="sxs-lookup"><span data-stu-id="58095-361">The Exchange server returned an error.</span></span> <span data-ttu-id="58095-362">詳細については、「diagnostics オブジェクト」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="58095-362">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="58095-363">ユーザーがネットワークに接続されていません。</span><span class="sxs-lookup"><span data-stu-id="58095-363">The user is no longer connected to the network.</span></span> <span data-ttu-id="58095-364">ネットワーク接続を確認し、もう一度実行してください。</span><span class="sxs-lookup"><span data-stu-id="58095-364">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="58095-365">要件</span><span class="sxs-lookup"><span data-stu-id="58095-365">Requirements</span></span>

|<span data-ttu-id="58095-366">要件</span><span class="sxs-lookup"><span data-stu-id="58095-366">Requirement</span></span>| <span data-ttu-id="58095-367">値</span><span class="sxs-lookup"><span data-stu-id="58095-367">Value</span></span>|
|---|---|
|[<span data-ttu-id="58095-368">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="58095-368">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="58095-369">1.0</span><span class="sxs-lookup"><span data-stu-id="58095-369">1.0</span></span>|
|[<span data-ttu-id="58095-370">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="58095-370">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="58095-371">ReadItem</span><span class="sxs-lookup"><span data-stu-id="58095-371">ReadItem</span></span>|
|[<span data-ttu-id="58095-372">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="58095-372">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="58095-373">新規作成と閲覧</span><span class="sxs-lookup"><span data-stu-id="58095-373">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="58095-374">例</span><span class="sxs-lookup"><span data-stu-id="58095-374">Example</span></span>

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

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="58095-375">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="58095-375">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="58095-376">ユーザーと Office アドインを識別するトークンを取得します。</span><span class="sxs-lookup"><span data-stu-id="58095-376">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="58095-377">`getUserIdentityTokenAsync` メソッドは、[アドインとユーザーをサード パーティのシステムで識別して認証](/outlook/add-ins/authentication)することのできるトークンを返します。</span><span class="sxs-lookup"><span data-stu-id="58095-377">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="58095-378">パラメーター</span><span class="sxs-lookup"><span data-stu-id="58095-378">Parameters</span></span>

|<span data-ttu-id="58095-379">名前</span><span class="sxs-lookup"><span data-stu-id="58095-379">Name</span></span>| <span data-ttu-id="58095-380">型</span><span class="sxs-lookup"><span data-stu-id="58095-380">Type</span></span>| <span data-ttu-id="58095-381">属性</span><span class="sxs-lookup"><span data-stu-id="58095-381">Attributes</span></span>| <span data-ttu-id="58095-382">説明</span><span class="sxs-lookup"><span data-stu-id="58095-382">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="58095-383">function</span><span class="sxs-lookup"><span data-stu-id="58095-383">function</span></span>||<span data-ttu-id="58095-384">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="58095-384">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="58095-385">トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="58095-385">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="58095-386">エラーが発生した場合は`asyncResult.error` 、 `asyncResult.diagnostics`プロパティとプロパティによって追加情報が提供されることがあります。</span><span class="sxs-lookup"><span data-stu-id="58095-386">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="58095-387">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="58095-387">Object</span></span>| <span data-ttu-id="58095-388">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="58095-388">&lt;optional&gt;</span></span>|<span data-ttu-id="58095-389">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="58095-389">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="58095-390">エラー</span><span class="sxs-lookup"><span data-stu-id="58095-390">Errors</span></span>

|<span data-ttu-id="58095-391">エラー コード</span><span class="sxs-lookup"><span data-stu-id="58095-391">Error code</span></span>|<span data-ttu-id="58095-392">説明</span><span class="sxs-lookup"><span data-stu-id="58095-392">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="58095-393">要求が失敗しました。</span><span class="sxs-lookup"><span data-stu-id="58095-393">The request has failed.</span></span> <span data-ttu-id="58095-394">HTTP エラーコードについては、diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="58095-394">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="58095-395">Exchange サーバーがエラーを返しました。</span><span class="sxs-lookup"><span data-stu-id="58095-395">The Exchange server returned an error.</span></span> <span data-ttu-id="58095-396">詳細については、「diagnostics オブジェクト」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="58095-396">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="58095-397">ユーザーがネットワークに接続されていません。</span><span class="sxs-lookup"><span data-stu-id="58095-397">The user is no longer connected to the network.</span></span> <span data-ttu-id="58095-398">ネットワーク接続を確認し、もう一度実行してください。</span><span class="sxs-lookup"><span data-stu-id="58095-398">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="58095-399">要件</span><span class="sxs-lookup"><span data-stu-id="58095-399">Requirements</span></span>

|<span data-ttu-id="58095-400">要件</span><span class="sxs-lookup"><span data-stu-id="58095-400">Requirement</span></span>| <span data-ttu-id="58095-401">値</span><span class="sxs-lookup"><span data-stu-id="58095-401">Value</span></span>|
|---|---|
|[<span data-ttu-id="58095-402">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="58095-402">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="58095-403">1.0</span><span class="sxs-lookup"><span data-stu-id="58095-403">1.0</span></span>|
|[<span data-ttu-id="58095-404">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="58095-404">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="58095-405">ReadItem</span><span class="sxs-lookup"><span data-stu-id="58095-405">ReadItem</span></span>|
|[<span data-ttu-id="58095-406">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="58095-406">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="58095-407">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="58095-407">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="58095-408">例</span><span class="sxs-lookup"><span data-stu-id="58095-408">Example</span></span>

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

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="58095-409">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="58095-409">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="58095-410">ユーザーのメールボックスをホストしている Exchange サーバー上の Exchange Web サービス (EWS) のサービスに対して非同期の要求を行います。</span><span class="sxs-lookup"><span data-stu-id="58095-410">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="58095-411">このメソッドは、次のシナリオではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="58095-411">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="58095-412">Outlook on iOS または Android</span><span class="sxs-lookup"><span data-stu-id="58095-412">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="58095-413">アドインが Gmail のメールボックスに読み込まれる場合</span><span class="sxs-lookup"><span data-stu-id="58095-413">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="58095-414">このような場合は、アドインでは [REST API を使用](/outlook/add-ins/use-rest-api)して、代わりにユーザーのメールボックスにアクセスする必要があります。</span><span class="sxs-lookup"><span data-stu-id="58095-414">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="58095-p128">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange. See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span><span class="sxs-lookup"><span data-stu-id="58095-p128">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange. See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="58095-417">`makeEwsRequestAsync` メソッドでは、フォルダー関連アイテムを要求できません。</span><span class="sxs-lookup"><span data-stu-id="58095-417">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="58095-418">XML 要求では UTF-8 エンコードを指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="58095-418">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="58095-p129">`makeEwsRequestAsync` メソッドを使用するには、アドインに **ReadWriteMailbox** アクセス許可が必要です。**ReadWriteMailbox** アクセス許可と、`makeEwsRequestAsync` メソッドで呼び出せる EWS 操作の使い方については、「[ユーザーのメールボックスへのメール アドイン アクセスのアクセス許可を指定する](/outlook/add-ins/understanding-outlook-add-in-permissions)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="58095-p129">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="58095-421">サーバー管理者は、クライアント アクセス サーバーの EWS ディレクトリで `OAuthAuthentication` を true に設定して、`makeEwsRequestAsync` メソッドで EWS 要求を行うことができるようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="58095-421">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="58095-422">バージョンの相違点</span><span class="sxs-lookup"><span data-stu-id="58095-422">Version differences</span></span>

<span data-ttu-id="58095-423">バージョン 15.0.4535.1004 より前のバージョンの Outlook で実行しているメール アプリで `makeEwsRequestAsync` メソッドを使う場合は、エンコード値を `ISO-8859-1` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="58095-423">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="58095-p130">Outlook on the web でメール アプリを実行している場合は、エンコード値を設定する必要はありません。mailbox.diagnostics.hostName プロパティを使って、メール アプリを Outlook で実行しているのか、Outlook on the web で実行しているのかを確認できます。mailbox.diagnostics.hostVersion プロパティを使って、どのバージョンの Outlook を使って実行しているかを確認できます。</span><span class="sxs-lookup"><span data-stu-id="58095-p130">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="58095-427">パラメーター</span><span class="sxs-lookup"><span data-stu-id="58095-427">Parameters</span></span>

|<span data-ttu-id="58095-428">名前</span><span class="sxs-lookup"><span data-stu-id="58095-428">Name</span></span>| <span data-ttu-id="58095-429">型</span><span class="sxs-lookup"><span data-stu-id="58095-429">Type</span></span>| <span data-ttu-id="58095-430">属性</span><span class="sxs-lookup"><span data-stu-id="58095-430">Attributes</span></span>| <span data-ttu-id="58095-431">説明</span><span class="sxs-lookup"><span data-stu-id="58095-431">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="58095-432">String</span><span class="sxs-lookup"><span data-stu-id="58095-432">String</span></span>||<span data-ttu-id="58095-433">EWS 要求です。</span><span class="sxs-lookup"><span data-stu-id="58095-433">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="58095-434">function</span><span class="sxs-lookup"><span data-stu-id="58095-434">function</span></span>||<span data-ttu-id="58095-435">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="58095-435">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="58095-p131">The XML result of the EWS call is provided as a string in the `asyncResult.value` property. If the result exceeds 1 MB in size, an error message is returned instead.</span><span class="sxs-lookup"><span data-stu-id="58095-p131">The XML result of the EWS call is provided as a string in the `asyncResult.value` property. If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="58095-438">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="58095-438">Object</span></span>| <span data-ttu-id="58095-439">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="58095-439">&lt;optional&gt;</span></span>|<span data-ttu-id="58095-440">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="58095-440">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="58095-441">要件</span><span class="sxs-lookup"><span data-stu-id="58095-441">Requirements</span></span>

|<span data-ttu-id="58095-442">要件</span><span class="sxs-lookup"><span data-stu-id="58095-442">Requirement</span></span>| <span data-ttu-id="58095-443">値</span><span class="sxs-lookup"><span data-stu-id="58095-443">Value</span></span>|
|---|---|
|[<span data-ttu-id="58095-444">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="58095-444">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="58095-445">1.0</span><span class="sxs-lookup"><span data-stu-id="58095-445">1.0</span></span>|
|[<span data-ttu-id="58095-446">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="58095-446">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="58095-447">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="58095-447">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="58095-448">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="58095-448">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="58095-449">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="58095-449">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="58095-450">例</span><span class="sxs-lookup"><span data-stu-id="58095-450">Example</span></span>

<span data-ttu-id="58095-451">次の例は、`makeEwsRequestAsync` を呼び出し、`GetItem` 操作を使って項目の件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="58095-451">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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
