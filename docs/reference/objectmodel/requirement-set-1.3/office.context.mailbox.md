---
title: Office. メールボックス要件セット1.3
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 04c8e3be7531e83279283a3d07e0ce14e1b11938
ms.sourcegitcommit: 3f5d7f4794e3d3c8bc3a79fa05c54157613b9376
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/02/2019
ms.locfileid: "36064747"
---
# <a name="mailbox"></a><span data-ttu-id="aa656-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="aa656-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="aa656-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="aa656-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="aa656-104">Microsoft Outlook の Outlook アドインオブジェクトモデルへのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="aa656-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="aa656-105">要件</span><span class="sxs-lookup"><span data-stu-id="aa656-105">Requirements</span></span>

|<span data-ttu-id="aa656-106">要件</span><span class="sxs-lookup"><span data-stu-id="aa656-106">Requirement</span></span>| <span data-ttu-id="aa656-107">値</span><span class="sxs-lookup"><span data-stu-id="aa656-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa656-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="aa656-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa656-109">1.0</span><span class="sxs-lookup"><span data-stu-id="aa656-109">1.0</span></span>|
|[<span data-ttu-id="aa656-110">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="aa656-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa656-111">制限あり</span><span class="sxs-lookup"><span data-stu-id="aa656-111">Restricted</span></span>|
|[<span data-ttu-id="aa656-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="aa656-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="aa656-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="aa656-113">Compose or Read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="aa656-114">名前空間</span><span class="sxs-lookup"><span data-stu-id="aa656-114">Namespaces</span></span>

<span data-ttu-id="aa656-115">[diagnostics](Office.context.mailbox.diagnostics.md):Outlook アドインに診断情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="aa656-115">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="aa656-116">[item](Office.context.mailbox.item.md):Outlook アドインのメッセージや予定にアクセスするためのメソッドとプロパティを提供します。</span><span class="sxs-lookup"><span data-stu-id="aa656-116">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="aa656-117">[userProfile](Office.context.mailbox.userProfile.md):Outlook アドインのユーザーに関する情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="aa656-117">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="aa656-118">メンバー</span><span class="sxs-lookup"><span data-stu-id="aa656-118">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="aa656-119">ewsUrl: String</span><span class="sxs-lookup"><span data-stu-id="aa656-119">ewsUrl: String</span></span>

<span data-ttu-id="aa656-120">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account.</span><span class="sxs-lookup"><span data-stu-id="aa656-120">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account.</span></span> <span data-ttu-id="aa656-121">Read mode only.</span><span class="sxs-lookup"><span data-stu-id="aa656-121">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="aa656-122">このメンバーは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="aa656-122">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="aa656-p102">
  `ewsUrl\` 値は、リモート サービスで、ユーザーのメールボックスに EWS 呼び出しを行うために使うことができます。たとえば、[選択したアイテムから添付ファイルを取得する](/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="aa656-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="aa656-125">アプリが閲覧モードで `ewsUrl` メンバーを呼び出すには、アプリのマニフェスト内に **ReadItem** アクセス許可が指定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="aa656-125">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="aa656-p103">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出してから、`ewsUrl` メンバーを使用する必要があります。アプリには、`saveAsync` メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="aa656-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="aa656-128">型</span><span class="sxs-lookup"><span data-stu-id="aa656-128">Type</span></span>

*   <span data-ttu-id="aa656-129">String</span><span class="sxs-lookup"><span data-stu-id="aa656-129">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="aa656-130">要件</span><span class="sxs-lookup"><span data-stu-id="aa656-130">Requirements</span></span>

|<span data-ttu-id="aa656-131">要件</span><span class="sxs-lookup"><span data-stu-id="aa656-131">Requirement</span></span>| <span data-ttu-id="aa656-132">値</span><span class="sxs-lookup"><span data-stu-id="aa656-132">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa656-133">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="aa656-133">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa656-134">1.0</span><span class="sxs-lookup"><span data-stu-id="aa656-134">1.0</span></span>|
|[<span data-ttu-id="aa656-135">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="aa656-135">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa656-136">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa656-136">ReadItem</span></span>|
|[<span data-ttu-id="aa656-137">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="aa656-137">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="aa656-138">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="aa656-138">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="aa656-139">メソッド</span><span class="sxs-lookup"><span data-stu-id="aa656-139">Methods</span></span>

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="aa656-140">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="aa656-140">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="aa656-141">REST 形式のアイテム ID を EWS 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="aa656-141">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="aa656-142">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="aa656-142">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="aa656-p104">REST API ([Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) や [Microsoft Graph](https://graph.microsoft.io/) など) で取得されたアイテム ID は、Exchange Web サービス (EWS) に使用される形式とは異なる形式を使用します。`convertToEwsId` メソッドは、REST 形式の ID を EWS 用の適切な形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="aa656-p104">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="aa656-145">パラメーター</span><span class="sxs-lookup"><span data-stu-id="aa656-145">Parameters</span></span>

|<span data-ttu-id="aa656-146">名前</span><span class="sxs-lookup"><span data-stu-id="aa656-146">Name</span></span>| <span data-ttu-id="aa656-147">種類</span><span class="sxs-lookup"><span data-stu-id="aa656-147">Type</span></span>| <span data-ttu-id="aa656-148">説明</span><span class="sxs-lookup"><span data-stu-id="aa656-148">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="aa656-149">String</span><span class="sxs-lookup"><span data-stu-id="aa656-149">String</span></span>|<span data-ttu-id="aa656-150">Outlook REST API 形式のアイテム ID</span><span class="sxs-lookup"><span data-stu-id="aa656-150">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="aa656-151">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="aa656-151">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.3)|<span data-ttu-id="aa656-152">アイテム ID の取得に使用された Outlook REST API のバージョンを示す値。</span><span class="sxs-lookup"><span data-stu-id="aa656-152">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="aa656-153">要件</span><span class="sxs-lookup"><span data-stu-id="aa656-153">Requirements</span></span>

|<span data-ttu-id="aa656-154">要件</span><span class="sxs-lookup"><span data-stu-id="aa656-154">Requirement</span></span>| <span data-ttu-id="aa656-155">値</span><span class="sxs-lookup"><span data-stu-id="aa656-155">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa656-156">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="aa656-156">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa656-157">1.3</span><span class="sxs-lookup"><span data-stu-id="aa656-157">1.3</span></span>|
|[<span data-ttu-id="aa656-158">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="aa656-158">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa656-159">制限あり</span><span class="sxs-lookup"><span data-stu-id="aa656-159">Restricted</span></span>|
|[<span data-ttu-id="aa656-160">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="aa656-160">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="aa656-161">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="aa656-161">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="aa656-162">戻り値:</span><span class="sxs-lookup"><span data-stu-id="aa656-162">Returns:</span></span>

<span data-ttu-id="aa656-163">型:String</span><span class="sxs-lookup"><span data-stu-id="aa656-163">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="aa656-164">例</span><span class="sxs-lookup"><span data-stu-id="aa656-164">Example</span></span>

```javascript
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-13"></a><span data-ttu-id="aa656-165">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.3)}</span><span class="sxs-lookup"><span data-stu-id="aa656-165">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.3)}</span></span>

<span data-ttu-id="aa656-166">クライアントのローカル時間で時間情報が含まれている辞書を取得します。</span><span class="sxs-lookup"><span data-stu-id="aa656-166">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="aa656-167">デスクトップまたは web 上の Outlook 用メールアプリは、日付と時刻に異なるタイムゾーンを使用できます。</span><span class="sxs-lookup"><span data-stu-id="aa656-167">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times.</span></span> <span data-ttu-id="aa656-168">デスクトップ上の Outlook では、クライアントコンピューターのタイムゾーンが使用されます。Outlook on the web では、Exchange 管理センター (EAC) で設定されているタイムゾーンが使用されます。</span><span class="sxs-lookup"><span data-stu-id="aa656-168">Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC).</span></span> <span data-ttu-id="aa656-169">日付と時刻の値を処理して、ユーザーインターフェイスに表示される値が、ユーザーが期待するタイムゾーンに常に一致するようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="aa656-169">You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="aa656-170">デスクトップクライアント上の Outlook でメールアプリが実行されている`convertToLocalClientTime`場合、このメソッドは、クライアントコンピューターのタイムゾーンに設定された値を持つ dictionary オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="aa656-170">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone.</span></span> <span data-ttu-id="aa656-171">メールアプリが web 上の Outlook で実行されている`convertToLocalClientTime`場合、このメソッドは、EAC で指定されたタイムゾーンに設定された値を持つ dictionary オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="aa656-171">If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="aa656-172">パラメーター</span><span class="sxs-lookup"><span data-stu-id="aa656-172">Parameters</span></span>

|<span data-ttu-id="aa656-173">名前</span><span class="sxs-lookup"><span data-stu-id="aa656-173">Name</span></span>| <span data-ttu-id="aa656-174">種類</span><span class="sxs-lookup"><span data-stu-id="aa656-174">Type</span></span>| <span data-ttu-id="aa656-175">説明</span><span class="sxs-lookup"><span data-stu-id="aa656-175">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="aa656-176">日付</span><span class="sxs-lookup"><span data-stu-id="aa656-176">Date</span></span>|<span data-ttu-id="aa656-177">日付オブジェクト</span><span class="sxs-lookup"><span data-stu-id="aa656-177">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="aa656-178">要件</span><span class="sxs-lookup"><span data-stu-id="aa656-178">Requirements</span></span>

|<span data-ttu-id="aa656-179">要件</span><span class="sxs-lookup"><span data-stu-id="aa656-179">Requirement</span></span>| <span data-ttu-id="aa656-180">値</span><span class="sxs-lookup"><span data-stu-id="aa656-180">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa656-181">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="aa656-181">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa656-182">1.0</span><span class="sxs-lookup"><span data-stu-id="aa656-182">1.0</span></span>|
|[<span data-ttu-id="aa656-183">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="aa656-183">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa656-184">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa656-184">ReadItem</span></span>|
|[<span data-ttu-id="aa656-185">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="aa656-185">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="aa656-186">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="aa656-186">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="aa656-187">戻り値:</span><span class="sxs-lookup"><span data-stu-id="aa656-187">Returns:</span></span>

<span data-ttu-id="aa656-188">型:[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="aa656-188">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.3)</span></span>

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="aa656-189">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="aa656-189">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="aa656-190">EWS 形式のアイテム ID を REST 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="aa656-190">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="aa656-191">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="aa656-191">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="aa656-p107">EWS または `itemId` プロパティで取得されるアイテム ID は、REST API ([Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) や [Microsoft Graph](https://graph.microsoft.io/) など) に使用される形式とは異なる形式を使用します。`convertToRestId` メソッドは、EWS 形式の ID を REST 用の適切な形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="aa656-p107">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="aa656-194">パラメーター</span><span class="sxs-lookup"><span data-stu-id="aa656-194">Parameters</span></span>

|<span data-ttu-id="aa656-195">名前</span><span class="sxs-lookup"><span data-stu-id="aa656-195">Name</span></span>| <span data-ttu-id="aa656-196">種類</span><span class="sxs-lookup"><span data-stu-id="aa656-196">Type</span></span>| <span data-ttu-id="aa656-197">説明</span><span class="sxs-lookup"><span data-stu-id="aa656-197">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="aa656-198">String</span><span class="sxs-lookup"><span data-stu-id="aa656-198">String</span></span>|<span data-ttu-id="aa656-199">Exchange Web サービス (EWS) 形式のアイテム ID</span><span class="sxs-lookup"><span data-stu-id="aa656-199">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="aa656-200">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="aa656-200">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.3)|<span data-ttu-id="aa656-201">変換後の ID を使用する Outlook REST API のバージョンを示す値。</span><span class="sxs-lookup"><span data-stu-id="aa656-201">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="aa656-202">要件</span><span class="sxs-lookup"><span data-stu-id="aa656-202">Requirements</span></span>

|<span data-ttu-id="aa656-203">要件</span><span class="sxs-lookup"><span data-stu-id="aa656-203">Requirement</span></span>| <span data-ttu-id="aa656-204">値</span><span class="sxs-lookup"><span data-stu-id="aa656-204">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa656-205">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="aa656-205">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa656-206">1.3</span><span class="sxs-lookup"><span data-stu-id="aa656-206">1.3</span></span>|
|[<span data-ttu-id="aa656-207">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="aa656-207">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa656-208">制限あり</span><span class="sxs-lookup"><span data-stu-id="aa656-208">Restricted</span></span>|
|[<span data-ttu-id="aa656-209">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="aa656-209">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="aa656-210">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="aa656-210">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="aa656-211">戻り値:</span><span class="sxs-lookup"><span data-stu-id="aa656-211">Returns:</span></span>

<span data-ttu-id="aa656-212">型:String</span><span class="sxs-lookup"><span data-stu-id="aa656-212">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="aa656-213">例</span><span class="sxs-lookup"><span data-stu-id="aa656-213">Example</span></span>

```javascript
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="aa656-214">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="aa656-214">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="aa656-215">時間情報が含まれているディクショナリから日付オブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="aa656-215">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="aa656-216">`convertToUtcClientTime` メソッドは、ローカルの日付と時刻を含むディクショナリを、ローカルの日付と時刻の正しい値を持つ日付オブジェクトに変換します。</span><span class="sxs-lookup"><span data-stu-id="aa656-216">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="aa656-217">パラメーター</span><span class="sxs-lookup"><span data-stu-id="aa656-217">Parameters</span></span>

|<span data-ttu-id="aa656-218">名前</span><span class="sxs-lookup"><span data-stu-id="aa656-218">Name</span></span>| <span data-ttu-id="aa656-219">種類</span><span class="sxs-lookup"><span data-stu-id="aa656-219">Type</span></span>| <span data-ttu-id="aa656-220">説明</span><span class="sxs-lookup"><span data-stu-id="aa656-220">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="aa656-221">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="aa656-221">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.3)|<span data-ttu-id="aa656-222">変換するローカル時刻の値。</span><span class="sxs-lookup"><span data-stu-id="aa656-222">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="aa656-223">要件</span><span class="sxs-lookup"><span data-stu-id="aa656-223">Requirements</span></span>

|<span data-ttu-id="aa656-224">要件</span><span class="sxs-lookup"><span data-stu-id="aa656-224">Requirement</span></span>| <span data-ttu-id="aa656-225">値</span><span class="sxs-lookup"><span data-stu-id="aa656-225">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa656-226">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="aa656-226">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa656-227">1.0</span><span class="sxs-lookup"><span data-stu-id="aa656-227">1.0</span></span>|
|[<span data-ttu-id="aa656-228">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="aa656-228">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa656-229">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa656-229">ReadItem</span></span>|
|[<span data-ttu-id="aa656-230">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="aa656-230">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="aa656-231">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="aa656-231">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="aa656-232">戻り値:</span><span class="sxs-lookup"><span data-stu-id="aa656-232">Returns:</span></span>

<span data-ttu-id="aa656-233">時間が UTC で表現された日付オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="aa656-233">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="aa656-234">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="aa656-234">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="aa656-235">Date</span><span class="sxs-lookup"><span data-stu-id="aa656-235">Date</span></span></dd>

</dl>

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="aa656-236">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="aa656-236">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="aa656-237">既存の予定を表示します。</span><span class="sxs-lookup"><span data-stu-id="aa656-237">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="aa656-238">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="aa656-238">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="aa656-239">`displayAppointmentForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存の予定を開きます。</span><span class="sxs-lookup"><span data-stu-id="aa656-239">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="aa656-240">Outlook on the Mac では、このメソッドを使用して、定期的なアイテムの一部ではない単一の予定を表示したり、定期的なアイテムのマスター予定を表示したりすることはできませんが、一連のインスタンスを表示することはできません。</span><span class="sxs-lookup"><span data-stu-id="aa656-240">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series.</span></span> <span data-ttu-id="aa656-241">これは、Mac 上の Outlook では、定期的なアイテムのインスタンスのプロパティ (アイテム ID を含む) にアクセスできないためです。</span><span class="sxs-lookup"><span data-stu-id="aa656-241">This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="aa656-242">Web 上の Outlook では、フォームの本文が 32 KB 以下の文字である場合にのみ、このメソッドは指定されたフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="aa656-242">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="aa656-243">指定のアイテム識別子が既存の予定を表していない場合は、クライアント コンピューターまたはデバイスで空のウィンドウが開き、エラー メッセージは返されません。</span><span class="sxs-lookup"><span data-stu-id="aa656-243">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="aa656-244">パラメーター</span><span class="sxs-lookup"><span data-stu-id="aa656-244">Parameters</span></span>

|<span data-ttu-id="aa656-245">名前</span><span class="sxs-lookup"><span data-stu-id="aa656-245">Name</span></span>| <span data-ttu-id="aa656-246">種類</span><span class="sxs-lookup"><span data-stu-id="aa656-246">Type</span></span>| <span data-ttu-id="aa656-247">説明</span><span class="sxs-lookup"><span data-stu-id="aa656-247">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="aa656-248">String</span><span class="sxs-lookup"><span data-stu-id="aa656-248">String</span></span>|<span data-ttu-id="aa656-249">既存の予定の Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="aa656-249">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="aa656-250">要件</span><span class="sxs-lookup"><span data-stu-id="aa656-250">Requirements</span></span>

|<span data-ttu-id="aa656-251">要件</span><span class="sxs-lookup"><span data-stu-id="aa656-251">Requirement</span></span>| <span data-ttu-id="aa656-252">値</span><span class="sxs-lookup"><span data-stu-id="aa656-252">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa656-253">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="aa656-253">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa656-254">1.0</span><span class="sxs-lookup"><span data-stu-id="aa656-254">1.0</span></span>|
|[<span data-ttu-id="aa656-255">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="aa656-255">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa656-256">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa656-256">ReadItem</span></span>|
|[<span data-ttu-id="aa656-257">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="aa656-257">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="aa656-258">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="aa656-258">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="aa656-259">例</span><span class="sxs-lookup"><span data-stu-id="aa656-259">Example</span></span>

```javascript
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

#### <a name="displaymessageformitemid"></a><span data-ttu-id="aa656-260">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="aa656-260">displayMessageForm(itemId)</span></span>

<span data-ttu-id="aa656-261">既存のメッセージを表示します。</span><span class="sxs-lookup"><span data-stu-id="aa656-261">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="aa656-262">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="aa656-262">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="aa656-263">`displayMessageForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存のメッセージを開きます。</span><span class="sxs-lookup"><span data-stu-id="aa656-263">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="aa656-264">Web 上の Outlook では、フォームの本文が 32 KB の文字数以下の場合にのみ、このメソッドは指定されたフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="aa656-264">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="aa656-265">指定のアイテム識別子が既存のメッセージを表していない場合は、クライアント コンピューターにはメッセージは表示されず、エラー メッセージも返されません。</span><span class="sxs-lookup"><span data-stu-id="aa656-265">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="aa656-p109">予定を表す `itemId` を含む `displayMessageForm` を使用しないでください。既存の予定を表示するには、`displayAppointmentForm` メソッドを使用します。新しい予定を作成するフォームを表示するには、`displayNewAppointmentForm` メソッドを使用します。</span><span class="sxs-lookup"><span data-stu-id="aa656-p109">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="aa656-268">パラメーター</span><span class="sxs-lookup"><span data-stu-id="aa656-268">Parameters</span></span>

|<span data-ttu-id="aa656-269">名前</span><span class="sxs-lookup"><span data-stu-id="aa656-269">Name</span></span>| <span data-ttu-id="aa656-270">種類</span><span class="sxs-lookup"><span data-stu-id="aa656-270">Type</span></span>| <span data-ttu-id="aa656-271">説明</span><span class="sxs-lookup"><span data-stu-id="aa656-271">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="aa656-272">String</span><span class="sxs-lookup"><span data-stu-id="aa656-272">String</span></span>|<span data-ttu-id="aa656-273">既存のメッセージの Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="aa656-273">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="aa656-274">要件</span><span class="sxs-lookup"><span data-stu-id="aa656-274">Requirements</span></span>

|<span data-ttu-id="aa656-275">要件</span><span class="sxs-lookup"><span data-stu-id="aa656-275">Requirement</span></span>| <span data-ttu-id="aa656-276">値</span><span class="sxs-lookup"><span data-stu-id="aa656-276">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa656-277">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="aa656-277">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa656-278">1.0</span><span class="sxs-lookup"><span data-stu-id="aa656-278">1.0</span></span>|
|[<span data-ttu-id="aa656-279">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="aa656-279">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa656-280">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa656-280">ReadItem</span></span>|
|[<span data-ttu-id="aa656-281">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="aa656-281">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="aa656-282">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="aa656-282">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="aa656-283">例</span><span class="sxs-lookup"><span data-stu-id="aa656-283">Example</span></span>

```javascript
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="aa656-284">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="aa656-284">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="aa656-285">新しい予定を作成するためのフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="aa656-285">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="aa656-286">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="aa656-286">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="aa656-p110">`displayNewAppointmentForm` メソッドを使用すると、ユーザーが新しい予定または会議を作成できるフォームが開きます。パラメーターを指定すると、予定のフォーム フィールドにパラメーターの内容が自動的に設定されます。</span><span class="sxs-lookup"><span data-stu-id="aa656-p110">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="aa656-289">Outlook on the web およびモバイルデバイスでは、このメソッドは常に出席者フィールドを含むフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="aa656-289">In Outlook on the web and mobile devices, this method always displays a form with an attendees field.</span></span> <span data-ttu-id="aa656-290">入力引数として出席者を指定しないと、このメソッドにより **[保存]** ボタンのあるフォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="aa656-290">If you do not specify any attendees as input arguments, the method displays a form with a **Save** button.</span></span> <span data-ttu-id="aa656-291">出席者を指定した場合には、フォームにその出席者と **[送信]** ボタンが表示されます。</span><span class="sxs-lookup"><span data-stu-id="aa656-291">If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="aa656-p112">Outlook リッチ クライアントと Outlook RT で、`requiredAttendees`、`optionalAttendees`、または `resources` パラメーターに出席者またはリソースを指定し、このメソッドを実行すると、**[送信]** ボタンがある会議フォームが表示されます。受信者を指定せずにこのメソッドを実行すると、**[保存して閉じる]** ボタンがある予定フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="aa656-p112">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="aa656-294">パラメーターのいずれかが指定のサイズ制限を超える場合、または不明なパラメーター名が指定されている場合は、例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="aa656-294">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="aa656-295">パラメーター</span><span class="sxs-lookup"><span data-stu-id="aa656-295">Parameters</span></span>

|<span data-ttu-id="aa656-296">名前</span><span class="sxs-lookup"><span data-stu-id="aa656-296">Name</span></span>| <span data-ttu-id="aa656-297">種類</span><span class="sxs-lookup"><span data-stu-id="aa656-297">Type</span></span>| <span data-ttu-id="aa656-298">説明</span><span class="sxs-lookup"><span data-stu-id="aa656-298">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="aa656-299">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="aa656-299">Object</span></span> | <span data-ttu-id="aa656-300">新しい予定を記述するパラメーターのディクショナリ。</span><span class="sxs-lookup"><span data-stu-id="aa656-300">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="aa656-301">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)&gt;</span><span class="sxs-lookup"><span data-stu-id="aa656-301">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)&gt;</span></span> | <span data-ttu-id="aa656-p113">予定に必要な各出席者について、メール アドレスを含む文字列の配列、または `EmailAddressDetails` オブジェクトを含む配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="aa656-p113">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="aa656-304">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)&gt;</span><span class="sxs-lookup"><span data-stu-id="aa656-304">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)&gt;</span></span> | <span data-ttu-id="aa656-p114">予定の各任意出席者について、メール アドレスを含む文字列の配列、または `EmailAddressDetails` オブジェクトを含む配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="aa656-p114">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="aa656-307">日付</span><span class="sxs-lookup"><span data-stu-id="aa656-307">Date</span></span> | <span data-ttu-id="aa656-308">予定の開始日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="aa656-308">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="aa656-309">日付</span><span class="sxs-lookup"><span data-stu-id="aa656-309">Date</span></span> | <span data-ttu-id="aa656-310">予定の終了日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="aa656-310">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="aa656-311">String</span><span class="sxs-lookup"><span data-stu-id="aa656-311">String</span></span> | <span data-ttu-id="aa656-p115">予定の場所を含む文字列。文字列は最大 255 文字に制限されます。</span><span class="sxs-lookup"><span data-stu-id="aa656-p115">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="aa656-314">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="aa656-314">Array.&lt;String&gt;</span></span> | <span data-ttu-id="aa656-p116">予定に必要なリソースを含む文字列の配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="aa656-p116">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="aa656-317">String</span><span class="sxs-lookup"><span data-stu-id="aa656-317">String</span></span> | <span data-ttu-id="aa656-p117">予定の件名を含む文字列です。文字列は最大 255 文字に制限されます。</span><span class="sxs-lookup"><span data-stu-id="aa656-p117">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="aa656-320">String</span><span class="sxs-lookup"><span data-stu-id="aa656-320">String</span></span> | <span data-ttu-id="aa656-p118">予定の本文。本文の内容は、最大サイズが 32 KB に制限されます。</span><span class="sxs-lookup"><span data-stu-id="aa656-p118">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="aa656-323">要件</span><span class="sxs-lookup"><span data-stu-id="aa656-323">Requirements</span></span>

|<span data-ttu-id="aa656-324">要件</span><span class="sxs-lookup"><span data-stu-id="aa656-324">Requirement</span></span>| <span data-ttu-id="aa656-325">値</span><span class="sxs-lookup"><span data-stu-id="aa656-325">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa656-326">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="aa656-326">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa656-327">1.0</span><span class="sxs-lookup"><span data-stu-id="aa656-327">1.0</span></span>|
|[<span data-ttu-id="aa656-328">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="aa656-328">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa656-329">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa656-329">ReadItem</span></span>|
|[<span data-ttu-id="aa656-330">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="aa656-330">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="aa656-331">読み取り</span><span class="sxs-lookup"><span data-stu-id="aa656-331">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="aa656-332">例</span><span class="sxs-lookup"><span data-stu-id="aa656-332">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="aa656-333">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="aa656-333">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="aa656-334">Exchange Server から添付ファイルやアイテムを取得するために使うトークンを含む文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="aa656-334">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="aa656-p119">`getCallbackTokenAsync` メソッドは、ユーザーのメールボックスをホストする Exchange Server から不透明なトークンを取得する非同期の呼び出しを行います。コールバック トークンの有効期間は 5 分です。</span><span class="sxs-lookup"><span data-stu-id="aa656-p119">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="aa656-p120">トークンと予定の識別子またはアイテムの識別子をサードパーティ システムに渡すことができます。サードパーティ システムは、トークンをベアラー承認トークンとして使用し、Exchange Web サービス (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) または [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) 操作を呼び出して、添付ファイルまたはアイテムを返します。たとえば、[選択したアイテムから添付ファイルを取得する](/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="aa656-p120">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="aa656-340">アプリが閲覧モードで `getCallbackTokenAsync` メソッドを呼び出すには、アプリのマニフェスト内に **ReadItem** アクセス許可が指定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="aa656-340">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="aa656-p121">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出してアイテムの識別子を `getCallbackTokenAsync` メソッドに渡す必要があります。アプリには、`saveAsync` メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="aa656-p121">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="aa656-343">パラメーター</span><span class="sxs-lookup"><span data-stu-id="aa656-343">Parameters</span></span>

|<span data-ttu-id="aa656-344">名前</span><span class="sxs-lookup"><span data-stu-id="aa656-344">Name</span></span>| <span data-ttu-id="aa656-345">型</span><span class="sxs-lookup"><span data-stu-id="aa656-345">Type</span></span>| <span data-ttu-id="aa656-346">属性</span><span class="sxs-lookup"><span data-stu-id="aa656-346">Attributes</span></span>| <span data-ttu-id="aa656-347">説明</span><span class="sxs-lookup"><span data-stu-id="aa656-347">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="aa656-348">function</span><span class="sxs-lookup"><span data-stu-id="aa656-348">function</span></span>||<span data-ttu-id="aa656-p122">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="aa656-p122">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="aa656-351">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="aa656-351">Object</span></span>| <span data-ttu-id="aa656-352">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="aa656-352">&lt;optional&gt;</span></span>|<span data-ttu-id="aa656-353">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="aa656-353">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="aa656-354">要件</span><span class="sxs-lookup"><span data-stu-id="aa656-354">Requirements</span></span>

|<span data-ttu-id="aa656-355">要件</span><span class="sxs-lookup"><span data-stu-id="aa656-355">Requirement</span></span>| <span data-ttu-id="aa656-356">値</span><span class="sxs-lookup"><span data-stu-id="aa656-356">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa656-357">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="aa656-357">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa656-358">1.3</span><span class="sxs-lookup"><span data-stu-id="aa656-358">1.3</span></span>|
|[<span data-ttu-id="aa656-359">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="aa656-359">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa656-360">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa656-360">ReadItem</span></span>|
|[<span data-ttu-id="aa656-361">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="aa656-361">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="aa656-362">新規作成と閲覧</span><span class="sxs-lookup"><span data-stu-id="aa656-362">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="aa656-363">例</span><span class="sxs-lookup"><span data-stu-id="aa656-363">Example</span></span>

```javascript
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="aa656-364">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="aa656-364">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="aa656-365">ユーザーと Office アドインを識別するトークンを取得します。</span><span class="sxs-lookup"><span data-stu-id="aa656-365">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="aa656-366">`getUserIdentityTokenAsync` メソッドは、[アドインとユーザーをサード パーティのシステムで識別して認証](/outlook/add-ins/authentication)することのできるトークンを返します。</span><span class="sxs-lookup"><span data-stu-id="aa656-366">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="aa656-367">パラメーター</span><span class="sxs-lookup"><span data-stu-id="aa656-367">Parameters</span></span>

|<span data-ttu-id="aa656-368">名前</span><span class="sxs-lookup"><span data-stu-id="aa656-368">Name</span></span>| <span data-ttu-id="aa656-369">型</span><span class="sxs-lookup"><span data-stu-id="aa656-369">Type</span></span>| <span data-ttu-id="aa656-370">属性</span><span class="sxs-lookup"><span data-stu-id="aa656-370">Attributes</span></span>| <span data-ttu-id="aa656-371">説明</span><span class="sxs-lookup"><span data-stu-id="aa656-371">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="aa656-372">関数</span><span class="sxs-lookup"><span data-stu-id="aa656-372">function</span></span>||<span data-ttu-id="aa656-373">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="aa656-373">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="aa656-374">トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="aa656-374">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="aa656-375">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="aa656-375">Object</span></span>| <span data-ttu-id="aa656-376">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="aa656-376">&lt;optional&gt;</span></span>|<span data-ttu-id="aa656-377">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="aa656-377">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="aa656-378">要件</span><span class="sxs-lookup"><span data-stu-id="aa656-378">Requirements</span></span>

|<span data-ttu-id="aa656-379">要件</span><span class="sxs-lookup"><span data-stu-id="aa656-379">Requirement</span></span>| <span data-ttu-id="aa656-380">値</span><span class="sxs-lookup"><span data-stu-id="aa656-380">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa656-381">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="aa656-381">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa656-382">1.0</span><span class="sxs-lookup"><span data-stu-id="aa656-382">1.0</span></span>|
|[<span data-ttu-id="aa656-383">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="aa656-383">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa656-384">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa656-384">ReadItem</span></span>|
|[<span data-ttu-id="aa656-385">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="aa656-385">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="aa656-386">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="aa656-386">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="aa656-387">例</span><span class="sxs-lookup"><span data-stu-id="aa656-387">Example</span></span>

```javascript
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="aa656-388">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="aa656-388">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="aa656-389">ユーザーのメールボックスをホストしている Exchange サーバー上の Exchange Web サービス (EWS) のサービスに対して非同期の要求を行います。</span><span class="sxs-lookup"><span data-stu-id="aa656-389">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="aa656-390">このメソッドは、次のシナリオではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="aa656-390">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="aa656-391">Outlook on iOS または Android</span><span class="sxs-lookup"><span data-stu-id="aa656-391">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="aa656-392">アドインが Gmail のメールボックスに読み込まれる場合</span><span class="sxs-lookup"><span data-stu-id="aa656-392">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="aa656-393">このような場合は、アドインでは [REST API を使用](/outlook/add-ins/use-rest-api)して、代わりにユーザーのメールボックスにアクセスする必要があります。</span><span class="sxs-lookup"><span data-stu-id="aa656-393">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="aa656-p123">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange. See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span><span class="sxs-lookup"><span data-stu-id="aa656-p123">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange. See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="aa656-396">`makeEwsRequestAsync` メソッドでは、フォルダー関連アイテムを要求できません。</span><span class="sxs-lookup"><span data-stu-id="aa656-396">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="aa656-397">XML 要求では UTF-8 エンコードを指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="aa656-397">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="aa656-p124">`makeEwsRequestAsync` メソッドを使用するには、アドインに **ReadWriteMailbox** アクセス許可が必要です。**ReadWriteMailbox** アクセス許可と、`makeEwsRequestAsync` メソッドで呼び出せる EWS 操作の使い方については、「[ユーザーのメールボックスへのメール アドイン アクセスのアクセス許可を指定する](/outlook/add-ins/understanding-outlook-add-in-permissions)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="aa656-p124">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="aa656-400">サーバー管理者は、クライアント アクセス サーバーの EWS ディレクトリで `OAuthAuthentication` を true に設定して、`makeEwsRequestAsync` メソッドで EWS 要求を行うことができるようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="aa656-400">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="aa656-401">バージョンの相違点</span><span class="sxs-lookup"><span data-stu-id="aa656-401">Version differences</span></span>

<span data-ttu-id="aa656-402">バージョン 15.0.4535.1004 より前のバージョンの Outlook で実行しているメール アプリで `makeEwsRequestAsync` メソッドを使う場合は、エンコード値を `ISO-8859-1` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="aa656-402">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="aa656-p125">Outlook on the web でメール アプリを実行している場合は、エンコード値を設定する必要はありません。mailbox.diagnostics.hostName プロパティを使って、メール アプリを Outlook で実行しているのか、Outlook on the web で実行しているのかを確認できます。mailbox.diagnostics.hostVersion プロパティを使って、どのバージョンの Outlook を使って実行しているかを確認できます。</span><span class="sxs-lookup"><span data-stu-id="aa656-p125">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="aa656-406">パラメーター</span><span class="sxs-lookup"><span data-stu-id="aa656-406">Parameters</span></span>

|<span data-ttu-id="aa656-407">名前</span><span class="sxs-lookup"><span data-stu-id="aa656-407">Name</span></span>| <span data-ttu-id="aa656-408">型</span><span class="sxs-lookup"><span data-stu-id="aa656-408">Type</span></span>| <span data-ttu-id="aa656-409">属性</span><span class="sxs-lookup"><span data-stu-id="aa656-409">Attributes</span></span>| <span data-ttu-id="aa656-410">説明</span><span class="sxs-lookup"><span data-stu-id="aa656-410">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="aa656-411">String</span><span class="sxs-lookup"><span data-stu-id="aa656-411">String</span></span>||<span data-ttu-id="aa656-412">EWS 要求です。</span><span class="sxs-lookup"><span data-stu-id="aa656-412">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="aa656-413">関数</span><span class="sxs-lookup"><span data-stu-id="aa656-413">function</span></span>||<span data-ttu-id="aa656-414">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="aa656-414">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="aa656-p126">The XML result of the EWS call is provided as a string in the `asyncResult.value` property. If the result exceeds 1 MB in size, an error message is returned instead.</span><span class="sxs-lookup"><span data-stu-id="aa656-p126">The XML result of the EWS call is provided as a string in the `asyncResult.value` property. If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="aa656-417">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="aa656-417">Object</span></span>| <span data-ttu-id="aa656-418">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="aa656-418">&lt;optional&gt;</span></span>|<span data-ttu-id="aa656-419">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="aa656-419">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="aa656-420">要件</span><span class="sxs-lookup"><span data-stu-id="aa656-420">Requirements</span></span>

|<span data-ttu-id="aa656-421">要件</span><span class="sxs-lookup"><span data-stu-id="aa656-421">Requirement</span></span>| <span data-ttu-id="aa656-422">値</span><span class="sxs-lookup"><span data-stu-id="aa656-422">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa656-423">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="aa656-423">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa656-424">1.0</span><span class="sxs-lookup"><span data-stu-id="aa656-424">1.0</span></span>|
|[<span data-ttu-id="aa656-425">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="aa656-425">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa656-426">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="aa656-426">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="aa656-427">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="aa656-427">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="aa656-428">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="aa656-428">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="aa656-429">例</span><span class="sxs-lookup"><span data-stu-id="aa656-429">Example</span></span>

<span data-ttu-id="aa656-430">次の例は、`makeEwsRequestAsync` を呼び出し、`GetItem` 操作を使って項目の件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="aa656-430">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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
