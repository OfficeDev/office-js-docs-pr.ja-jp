---
title: Office.context.mailbox - 要件セット 1.3
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: edf9bcaa842660007bea85b497308aa7949f8085
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/16/2019
ms.locfileid: "30068225"
---
# <a name="mailbox"></a><span data-ttu-id="e684c-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="e684c-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="e684c-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="e684c-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="e684c-104">Microsoft Outlook と Microsoft Outlook on the web の Outlook アドイン オブジェクト モデルへのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="e684c-104">Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="e684c-105">要件</span><span class="sxs-lookup"><span data-stu-id="e684c-105">Requirements</span></span>

|<span data-ttu-id="e684c-106">要件</span><span class="sxs-lookup"><span data-stu-id="e684c-106">Requirement</span></span>| <span data-ttu-id="e684c-107">値</span><span class="sxs-lookup"><span data-stu-id="e684c-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="e684c-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e684c-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e684c-109">1.0</span><span class="sxs-lookup"><span data-stu-id="e684c-109">1.0</span></span>|
|[<span data-ttu-id="e684c-110">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e684c-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e684c-111">制限あり</span><span class="sxs-lookup"><span data-stu-id="e684c-111">Restricted</span></span>|
|[<span data-ttu-id="e684c-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e684c-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e684c-113">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="e684c-113">Compose or Read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="e684c-114">名前空間</span><span class="sxs-lookup"><span data-stu-id="e684c-114">Namespaces</span></span>

<span data-ttu-id="e684c-115">[diagnostics](Office.context.mailbox.diagnostics.md):Outlook アドインに診断情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="e684c-115">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="e684c-116">[item](Office.context.mailbox.item.md):Outlook アドインのメッセージや予定にアクセスするためのメソッドとプロパティを提供します。</span><span class="sxs-lookup"><span data-stu-id="e684c-116">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="e684c-117">[userProfile](Office.context.mailbox.userProfile.md):Outlook アドインのユーザーに関する情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="e684c-117">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="e684c-118">メンバー</span><span class="sxs-lookup"><span data-stu-id="e684c-118">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="e684c-119">ewsUrl :String</span><span class="sxs-lookup"><span data-stu-id="e684c-119">ewsUrl :String</span></span>

<span data-ttu-id="e684c-p101">このメール アカウントの Exchange Web サービス (EWS) エンドポイントの URL を取得します。読み取りモードのみです。</span><span class="sxs-lookup"><span data-stu-id="e684c-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="e684c-122">このメンバーは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="e684c-122">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="e684c-p102">
  `ewsUrl\` 値は、リモート サービスで、ユーザーのメールボックスに EWS 呼び出しを行うために使うことができます。たとえば、[選択したアイテムから添付ファイルを取得する](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="e684c-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="e684c-125">アプリが閲覧モードで `ewsUrl` メンバーを呼び出すには、アプリのマニフェスト内に **ReadItem** アクセス許可が指定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="e684c-125">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="e684c-p103">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出してから、`ewsUrl` メンバーを使用する必要があります。アプリには、`saveAsync` メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="e684c-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="e684c-128">型</span><span class="sxs-lookup"><span data-stu-id="e684c-128">Type</span></span>

*   <span data-ttu-id="e684c-129">String</span><span class="sxs-lookup"><span data-stu-id="e684c-129">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e684c-130">Requirements</span><span class="sxs-lookup"><span data-stu-id="e684c-130">Requirements</span></span>

|<span data-ttu-id="e684c-131">要件</span><span class="sxs-lookup"><span data-stu-id="e684c-131">Requirement</span></span>| <span data-ttu-id="e684c-132">値</span><span class="sxs-lookup"><span data-stu-id="e684c-132">Value</span></span>|
|---|---|
|[<span data-ttu-id="e684c-133">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e684c-133">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e684c-134">1.0</span><span class="sxs-lookup"><span data-stu-id="e684c-134">1.0</span></span>|
|[<span data-ttu-id="e684c-135">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e684c-135">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e684c-136">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e684c-136">ReadItem</span></span>|
|[<span data-ttu-id="e684c-137">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e684c-137">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e684c-138">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="e684c-138">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="e684c-139">メソッド</span><span class="sxs-lookup"><span data-stu-id="e684c-139">Methods</span></span>

####  <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="e684c-140">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="e684c-140">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="e684c-141">REST 形式のアイテム ID を EWS 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="e684c-141">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="e684c-142">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="e684c-142">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="e684c-p104">REST API ([Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) や [Microsoft Graph](https://graph.microsoft.io/) など) で取得されたアイテム ID は、Exchange Web サービス (EWS) に使用される形式とは異なる形式を使用します。`convertToEwsId` メソッドは、REST 形式の ID を EWS 用の適切な形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="e684c-p104">Item IDs retrieved via a REST API (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e684c-145">パラメーター</span><span class="sxs-lookup"><span data-stu-id="e684c-145">Parameters</span></span>

|<span data-ttu-id="e684c-146">名前</span><span class="sxs-lookup"><span data-stu-id="e684c-146">Name</span></span>| <span data-ttu-id="e684c-147">型</span><span class="sxs-lookup"><span data-stu-id="e684c-147">Type</span></span>| <span data-ttu-id="e684c-148">説明</span><span class="sxs-lookup"><span data-stu-id="e684c-148">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="e684c-149">String</span><span class="sxs-lookup"><span data-stu-id="e684c-149">String</span></span>|<span data-ttu-id="e684c-150">Outlook REST API 形式のアイテム ID</span><span class="sxs-lookup"><span data-stu-id="e684c-150">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="e684c-151">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="e684c-151">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_3/office.mailboxenums.restversion)|<span data-ttu-id="e684c-152">アイテム ID の取得に使用された Outlook REST API のバージョンを示す値。</span><span class="sxs-lookup"><span data-stu-id="e684c-152">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e684c-153">要件</span><span class="sxs-lookup"><span data-stu-id="e684c-153">Requirements</span></span>

|<span data-ttu-id="e684c-154">要件</span><span class="sxs-lookup"><span data-stu-id="e684c-154">Requirement</span></span>| <span data-ttu-id="e684c-155">値</span><span class="sxs-lookup"><span data-stu-id="e684c-155">Value</span></span>|
|---|---|
|[<span data-ttu-id="e684c-156">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e684c-156">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e684c-157">1.3</span><span class="sxs-lookup"><span data-stu-id="e684c-157">1.3</span></span>|
|[<span data-ttu-id="e684c-158">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e684c-158">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e684c-159">制限あり</span><span class="sxs-lookup"><span data-stu-id="e684c-159">Restricted</span></span>|
|[<span data-ttu-id="e684c-160">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e684c-160">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e684c-161">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="e684c-161">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e684c-162">戻り値:</span><span class="sxs-lookup"><span data-stu-id="e684c-162">Returns:</span></span>

<span data-ttu-id="e684c-163">型:String</span><span class="sxs-lookup"><span data-stu-id="e684c-163">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="e684c-164">例</span><span class="sxs-lookup"><span data-stu-id="e684c-164">Example</span></span>

```javascript
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook13officelocalclienttime"></a><span data-ttu-id="e684c-165">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_3/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="e684c-165">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_3/office.LocalClientTime)}</span></span>

<span data-ttu-id="e684c-166">クライアントのローカル時間で時間情報が含まれている辞書を取得します。</span><span class="sxs-lookup"><span data-stu-id="e684c-166">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="e684c-p105">Outlook 用メール アプリや Outlook Web App で使う日付と時刻では、異なるタイム ゾーンを使うことができます。Outlook では、クライアント コンピューターのタイム ゾーンを使います。Outlook Web App では、Exchange 管理センター (EAC) で設定されたタイム ゾーンを使います。ユーザー インターフェイスに表示される値が、常にユーザーが期待するタイム ゾーンと一致するように日付と時刻の値を処理する必要があります。</span><span class="sxs-lookup"><span data-stu-id="e684c-p105">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="e684c-p106">Outlook でメール アプリが実行されている場合、`convertToLocalClientTime` メソッドは、クライアント コンピューターのタイム ゾーンに設定された値のディクショナリ オブジェクトを返します。Outlook Web Apps でメール アプリが実行されている場合、`convertToLocalClientTime` メソッドは、EAC に指定されたタイム ゾーンに設定された値のディクショナリ オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="e684c-p106">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e684c-172">パラメーター</span><span class="sxs-lookup"><span data-stu-id="e684c-172">Parameters</span></span>

|<span data-ttu-id="e684c-173">名前</span><span class="sxs-lookup"><span data-stu-id="e684c-173">Name</span></span>| <span data-ttu-id="e684c-174">型</span><span class="sxs-lookup"><span data-stu-id="e684c-174">Type</span></span>| <span data-ttu-id="e684c-175">説明</span><span class="sxs-lookup"><span data-stu-id="e684c-175">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="e684c-176">Date</span><span class="sxs-lookup"><span data-stu-id="e684c-176">Date</span></span>|<span data-ttu-id="e684c-177">日付オブジェクト</span><span class="sxs-lookup"><span data-stu-id="e684c-177">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e684c-178">要件</span><span class="sxs-lookup"><span data-stu-id="e684c-178">Requirements</span></span>

|<span data-ttu-id="e684c-179">要件</span><span class="sxs-lookup"><span data-stu-id="e684c-179">Requirement</span></span>| <span data-ttu-id="e684c-180">値</span><span class="sxs-lookup"><span data-stu-id="e684c-180">Value</span></span>|
|---|---|
|[<span data-ttu-id="e684c-181">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e684c-181">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e684c-182">1.0</span><span class="sxs-lookup"><span data-stu-id="e684c-182">1.0</span></span>|
|[<span data-ttu-id="e684c-183">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e684c-183">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e684c-184">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e684c-184">ReadItem</span></span>|
|[<span data-ttu-id="e684c-185">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e684c-185">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e684c-186">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="e684c-186">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e684c-187">戻り値:</span><span class="sxs-lookup"><span data-stu-id="e684c-187">Returns:</span></span>

<span data-ttu-id="e684c-188">型:[LocalClientTime](/javascript/api/outlook_1_3/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="e684c-188">Type: [LocalClientTime](/javascript/api/outlook_1_3/office.LocalClientTime)</span></span>

####  <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="e684c-189">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="e684c-189">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="e684c-190">EWS 形式のアイテム ID を REST 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="e684c-190">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="e684c-191">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="e684c-191">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="e684c-p107">EWS または `itemId` プロパティで取得されるアイテム ID は、REST API ([Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) や [Microsoft Graph](https://graph.microsoft.io/) など) に使用される形式とは異なる形式を使用します。`convertToRestId` メソッドは、EWS 形式の ID を REST 用の適切な形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="e684c-p107">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e684c-194">パラメーター</span><span class="sxs-lookup"><span data-stu-id="e684c-194">Parameters</span></span>

|<span data-ttu-id="e684c-195">名前</span><span class="sxs-lookup"><span data-stu-id="e684c-195">Name</span></span>| <span data-ttu-id="e684c-196">型</span><span class="sxs-lookup"><span data-stu-id="e684c-196">Type</span></span>| <span data-ttu-id="e684c-197">説明</span><span class="sxs-lookup"><span data-stu-id="e684c-197">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="e684c-198">String</span><span class="sxs-lookup"><span data-stu-id="e684c-198">String</span></span>|<span data-ttu-id="e684c-199">Exchange Web サービス (EWS) 形式のアイテム ID</span><span class="sxs-lookup"><span data-stu-id="e684c-199">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="e684c-200">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="e684c-200">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_3/office.mailboxenums.restversion)|<span data-ttu-id="e684c-201">変換後の ID を使用する Outlook REST API のバージョンを示す値。</span><span class="sxs-lookup"><span data-stu-id="e684c-201">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e684c-202">要件</span><span class="sxs-lookup"><span data-stu-id="e684c-202">Requirements</span></span>

|<span data-ttu-id="e684c-203">要件</span><span class="sxs-lookup"><span data-stu-id="e684c-203">Requirement</span></span>| <span data-ttu-id="e684c-204">値</span><span class="sxs-lookup"><span data-stu-id="e684c-204">Value</span></span>|
|---|---|
|[<span data-ttu-id="e684c-205">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e684c-205">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e684c-206">1.3</span><span class="sxs-lookup"><span data-stu-id="e684c-206">1.3</span></span>|
|[<span data-ttu-id="e684c-207">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e684c-207">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e684c-208">制限あり</span><span class="sxs-lookup"><span data-stu-id="e684c-208">Restricted</span></span>|
|[<span data-ttu-id="e684c-209">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e684c-209">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e684c-210">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="e684c-210">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e684c-211">戻り値:</span><span class="sxs-lookup"><span data-stu-id="e684c-211">Returns:</span></span>

<span data-ttu-id="e684c-212">型:String</span><span class="sxs-lookup"><span data-stu-id="e684c-212">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="e684c-213">例</span><span class="sxs-lookup"><span data-stu-id="e684c-213">Example</span></span>

```javascript
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="e684c-214">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="e684c-214">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="e684c-215">時間情報が含まれているディクショナリから日付オブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="e684c-215">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="e684c-216">`convertToUtcClientTime` メソッドは、ローカルの日付と時刻を含むディクショナリを、ローカルの日付と時刻の正しい値を持つ日付オブジェクトに変換します。</span><span class="sxs-lookup"><span data-stu-id="e684c-216">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e684c-217">パラメーター</span><span class="sxs-lookup"><span data-stu-id="e684c-217">Parameters</span></span>

|<span data-ttu-id="e684c-218">名前</span><span class="sxs-lookup"><span data-stu-id="e684c-218">Name</span></span>| <span data-ttu-id="e684c-219">型</span><span class="sxs-lookup"><span data-stu-id="e684c-219">Type</span></span>| <span data-ttu-id="e684c-220">説明</span><span class="sxs-lookup"><span data-stu-id="e684c-220">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="e684c-221">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="e684c-221">LocalClientTime</span></span>](/javascript/api/outlook_1_3/office.LocalClientTime)|<span data-ttu-id="e684c-222">変換するローカル時刻の値。</span><span class="sxs-lookup"><span data-stu-id="e684c-222">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e684c-223">要件</span><span class="sxs-lookup"><span data-stu-id="e684c-223">Requirements</span></span>

|<span data-ttu-id="e684c-224">要件</span><span class="sxs-lookup"><span data-stu-id="e684c-224">Requirement</span></span>| <span data-ttu-id="e684c-225">値</span><span class="sxs-lookup"><span data-stu-id="e684c-225">Value</span></span>|
|---|---|
|[<span data-ttu-id="e684c-226">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e684c-226">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e684c-227">1.0</span><span class="sxs-lookup"><span data-stu-id="e684c-227">1.0</span></span>|
|[<span data-ttu-id="e684c-228">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e684c-228">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e684c-229">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e684c-229">ReadItem</span></span>|
|[<span data-ttu-id="e684c-230">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e684c-230">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e684c-231">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="e684c-231">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e684c-232">戻り値:</span><span class="sxs-lookup"><span data-stu-id="e684c-232">Returns:</span></span>

<span data-ttu-id="e684c-233">時間が UTC で表現された日付オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="e684c-233">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="e684c-234">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="e684c-234">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="e684c-235">Date</span><span class="sxs-lookup"><span data-stu-id="e684c-235">Date</span></span></dd>

</dl>

####  <a name="displayappointmentformitemid"></a><span data-ttu-id="e684c-236">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="e684c-236">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="e684c-237">既存の予定を表示します。</span><span class="sxs-lookup"><span data-stu-id="e684c-237">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="e684c-238">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="e684c-238">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="e684c-239">`displayAppointmentForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存の予定を開きます。</span><span class="sxs-lookup"><span data-stu-id="e684c-239">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="e684c-p108">Outlook for Mac では、この方法を使って、定期的な系列の一部ではない単一の予定や定期的な系列のマスター予定を表示できます。ただし、系列のインスタンスは表示できません。これは、Outlook for Mac においては定期的な系列のインスタンスのプロパティ (アイテム ID を含む) にアクセスできないためです。</span><span class="sxs-lookup"><span data-stu-id="e684c-p108">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="e684c-242">Outlook Web App では、このメソッドは指定されたフォームの本文が 32 KB 以下の文字数の場合にフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="e684c-242">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="e684c-243">指定のアイテム識別子が既存の予定を表していない場合は、クライアント コンピューターまたはデバイスで空のウィンドウが開き、エラー メッセージは返されません。</span><span class="sxs-lookup"><span data-stu-id="e684c-243">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e684c-244">パラメーター</span><span class="sxs-lookup"><span data-stu-id="e684c-244">Parameters</span></span>

|<span data-ttu-id="e684c-245">名前</span><span class="sxs-lookup"><span data-stu-id="e684c-245">Name</span></span>| <span data-ttu-id="e684c-246">種類</span><span class="sxs-lookup"><span data-stu-id="e684c-246">Type</span></span>| <span data-ttu-id="e684c-247">説明</span><span class="sxs-lookup"><span data-stu-id="e684c-247">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="e684c-248">String</span><span class="sxs-lookup"><span data-stu-id="e684c-248">String</span></span>|<span data-ttu-id="e684c-249">既存の予定の Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="e684c-249">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e684c-250">要件</span><span class="sxs-lookup"><span data-stu-id="e684c-250">Requirements</span></span>

|<span data-ttu-id="e684c-251">要件</span><span class="sxs-lookup"><span data-stu-id="e684c-251">Requirement</span></span>| <span data-ttu-id="e684c-252">値</span><span class="sxs-lookup"><span data-stu-id="e684c-252">Value</span></span>|
|---|---|
|[<span data-ttu-id="e684c-253">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e684c-253">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e684c-254">1.0</span><span class="sxs-lookup"><span data-stu-id="e684c-254">1.0</span></span>|
|[<span data-ttu-id="e684c-255">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e684c-255">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e684c-256">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e684c-256">ReadItem</span></span>|
|[<span data-ttu-id="e684c-257">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e684c-257">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e684c-258">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="e684c-258">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e684c-259">例</span><span class="sxs-lookup"><span data-stu-id="e684c-259">Example</span></span>

```javascript
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  <a name="displaymessageformitemid"></a><span data-ttu-id="e684c-260">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="e684c-260">displayMessageForm(itemId)</span></span>

<span data-ttu-id="e684c-261">既存のメッセージを表示します。</span><span class="sxs-lookup"><span data-stu-id="e684c-261">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="e684c-262">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="e684c-262">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="e684c-263">`displayMessageForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存のメッセージを開きます。</span><span class="sxs-lookup"><span data-stu-id="e684c-263">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="e684c-264">Outlook Web App では、このメソッドは指定されたフォームの本文が 32 KB 以下の文字数の場合にフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="e684c-264">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="e684c-265">指定のアイテム識別子が既存のメッセージを表していない場合は、クライアント コンピューターにはメッセージは表示されず、エラー メッセージも返されません。</span><span class="sxs-lookup"><span data-stu-id="e684c-265">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="e684c-p109">予定を表す `itemId` を含む `displayMessageForm` を使用しないでください。既存の予定を表示するには、`displayAppointmentForm` メソッドを使用します。新しい予定を作成するフォームを表示するには、`displayNewAppointmentForm` メソッドを使用します。</span><span class="sxs-lookup"><span data-stu-id="e684c-p109">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e684c-268">パラメーター</span><span class="sxs-lookup"><span data-stu-id="e684c-268">Parameters</span></span>

|<span data-ttu-id="e684c-269">名前</span><span class="sxs-lookup"><span data-stu-id="e684c-269">Name</span></span>| <span data-ttu-id="e684c-270">種類</span><span class="sxs-lookup"><span data-stu-id="e684c-270">Type</span></span>| <span data-ttu-id="e684c-271">説明</span><span class="sxs-lookup"><span data-stu-id="e684c-271">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="e684c-272">String</span><span class="sxs-lookup"><span data-stu-id="e684c-272">String</span></span>|<span data-ttu-id="e684c-273">既存のメッセージの Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="e684c-273">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e684c-274">要件</span><span class="sxs-lookup"><span data-stu-id="e684c-274">Requirements</span></span>

|<span data-ttu-id="e684c-275">要件</span><span class="sxs-lookup"><span data-stu-id="e684c-275">Requirement</span></span>| <span data-ttu-id="e684c-276">値</span><span class="sxs-lookup"><span data-stu-id="e684c-276">Value</span></span>|
|---|---|
|[<span data-ttu-id="e684c-277">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e684c-277">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e684c-278">1.0</span><span class="sxs-lookup"><span data-stu-id="e684c-278">1.0</span></span>|
|[<span data-ttu-id="e684c-279">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e684c-279">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e684c-280">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e684c-280">ReadItem</span></span>|
|[<span data-ttu-id="e684c-281">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e684c-281">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e684c-282">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="e684c-282">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e684c-283">例</span><span class="sxs-lookup"><span data-stu-id="e684c-283">Example</span></span>

```javascript
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="e684c-284">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="e684c-284">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="e684c-285">新しい予定を作成するためのフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="e684c-285">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="e684c-286">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="e684c-286">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="e684c-p110">`displayNewAppointmentForm` メソッドを使用すると、ユーザーが新しい予定または会議を作成できるフォームが開きます。パラメーターを指定すると、予定のフォーム フィールドにパラメーターの内容が自動的に設定されます。</span><span class="sxs-lookup"><span data-stu-id="e684c-p110">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="e684c-p111">このメソッドは、Outlook Web App と OWA for Devices において、出席者フィールドが含まれるフォームを必ず表示します。入力引数として出席者を指定しないと、このメソッドにより **[保存]** ボタンのあるフォームが表示されます。出席者を指定した場合には、フォームにその出席者と **[送信]** ボタンが表示されます。</span><span class="sxs-lookup"><span data-stu-id="e684c-p111">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="e684c-p112">Outlook リッチ クライアントと Outlook RT で、`requiredAttendees`、`optionalAttendees`、または `resources` パラメーターに出席者またはリソースを指定し、このメソッドを実行すると、**[送信]** ボタンがある会議フォームが表示されます。受信者を指定せずにこのメソッドを実行すると、**[保存して閉じる]** ボタンがある予定フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="e684c-p112">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="e684c-294">パラメーターのいずれかが指定のサイズ制限を超える場合、または不明なパラメーター名が指定されている場合は、例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="e684c-294">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e684c-295">パラメーター</span><span class="sxs-lookup"><span data-stu-id="e684c-295">Parameters</span></span>

|<span data-ttu-id="e684c-296">名前</span><span class="sxs-lookup"><span data-stu-id="e684c-296">Name</span></span>| <span data-ttu-id="e684c-297">種類</span><span class="sxs-lookup"><span data-stu-id="e684c-297">Type</span></span>| <span data-ttu-id="e684c-298">説明</span><span class="sxs-lookup"><span data-stu-id="e684c-298">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="e684c-299">Object</span><span class="sxs-lookup"><span data-stu-id="e684c-299">Object</span></span> | <span data-ttu-id="e684c-300">新しい予定を記述するパラメーターのディクショナリ。</span><span class="sxs-lookup"><span data-stu-id="e684c-300">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="e684c-301">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="e684c-301">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="e684c-p113">予定に必要な各出席者について、メール アドレスを含む文字列の配列、または `EmailAddressDetails` オブジェクトを含む配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="e684c-p113">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="e684c-304">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="e684c-304">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="e684c-p114">予定の各任意出席者について、メール アドレスを含む文字列の配列、または `EmailAddressDetails` オブジェクトを含む配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="e684c-p114">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="e684c-307">日付</span><span class="sxs-lookup"><span data-stu-id="e684c-307">Date</span></span> | <span data-ttu-id="e684c-308">予定の開始日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="e684c-308">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="e684c-309">Date</span><span class="sxs-lookup"><span data-stu-id="e684c-309">Date</span></span> | <span data-ttu-id="e684c-310">予定の終了日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="e684c-310">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="e684c-311">String</span><span class="sxs-lookup"><span data-stu-id="e684c-311">String</span></span> | <span data-ttu-id="e684c-p115">予定の場所を含む文字列。文字列は最大 255 文字に制限されます。</span><span class="sxs-lookup"><span data-stu-id="e684c-p115">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="e684c-314">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="e684c-314">Array.&lt;String&gt;</span></span> | <span data-ttu-id="e684c-p116">予定に必要なリソースを含む文字列の配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="e684c-p116">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="e684c-317">String</span><span class="sxs-lookup"><span data-stu-id="e684c-317">String</span></span> | <span data-ttu-id="e684c-p117">予定の件名を含む文字列です。文字列は最大 255 文字に制限されます。</span><span class="sxs-lookup"><span data-stu-id="e684c-p117">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="e684c-320">String</span><span class="sxs-lookup"><span data-stu-id="e684c-320">String</span></span> | <span data-ttu-id="e684c-p118">予定の本文。本文の内容は、最大サイズが 32 KB に制限されます。</span><span class="sxs-lookup"><span data-stu-id="e684c-p118">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="e684c-323">要件</span><span class="sxs-lookup"><span data-stu-id="e684c-323">Requirements</span></span>

|<span data-ttu-id="e684c-324">要件</span><span class="sxs-lookup"><span data-stu-id="e684c-324">Requirement</span></span>| <span data-ttu-id="e684c-325">値</span><span class="sxs-lookup"><span data-stu-id="e684c-325">Value</span></span>|
|---|---|
|[<span data-ttu-id="e684c-326">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e684c-326">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e684c-327">1.0</span><span class="sxs-lookup"><span data-stu-id="e684c-327">1.0</span></span>|
|[<span data-ttu-id="e684c-328">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e684c-328">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e684c-329">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e684c-329">ReadItem</span></span>|
|[<span data-ttu-id="e684c-330">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e684c-330">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e684c-331">読み取り</span><span class="sxs-lookup"><span data-stu-id="e684c-331">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e684c-332">例</span><span class="sxs-lookup"><span data-stu-id="e684c-332">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="e684c-333">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="e684c-333">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="e684c-334">Exchange Server から添付ファイルやアイテムを取得するために使うトークンを含む文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="e684c-334">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="e684c-p119">`getCallbackTokenAsync` メソッドは、ユーザーのメールボックスをホストする Exchange Server から不透明なトークンを取得する非同期の呼び出しを行います。コールバック トークンの有効期間は 5 分です。</span><span class="sxs-lookup"><span data-stu-id="e684c-p119">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="e684c-p120">トークンと予定の識別子またはアイテムの識別子をサードパーティ システムに渡すことができます。サードパーティ システムは、トークンをベアラー承認トークンとして使用し、Exchange Web サービス (EWS) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) または [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) 操作を呼び出して、添付ファイルまたはアイテムを返します。たとえば、[選択したアイテムから添付ファイルを取得する](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="e684c-p120">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="e684c-340">アプリが閲覧モードで `getCallbackTokenAsync` メソッドを呼び出すには、アプリのマニフェスト内に **ReadItem** アクセス許可が指定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="e684c-340">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="e684c-p121">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出してアイテムの識別子を `getCallbackTokenAsync` メソッドに渡す必要があります。アプリには、`saveAsync` メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="e684c-p121">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e684c-343">パラメーター</span><span class="sxs-lookup"><span data-stu-id="e684c-343">Parameters</span></span>

|<span data-ttu-id="e684c-344">名前</span><span class="sxs-lookup"><span data-stu-id="e684c-344">Name</span></span>| <span data-ttu-id="e684c-345">型</span><span class="sxs-lookup"><span data-stu-id="e684c-345">Type</span></span>| <span data-ttu-id="e684c-346">属性</span><span class="sxs-lookup"><span data-stu-id="e684c-346">Attributes</span></span>| <span data-ttu-id="e684c-347">説明</span><span class="sxs-lookup"><span data-stu-id="e684c-347">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="e684c-348">function</span><span class="sxs-lookup"><span data-stu-id="e684c-348">function</span></span>||<span data-ttu-id="e684c-p122">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="e684c-p122">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="e684c-351">Object</span><span class="sxs-lookup"><span data-stu-id="e684c-351">Object</span></span>| <span data-ttu-id="e684c-352">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="e684c-352">&lt;optional&gt;</span></span>|<span data-ttu-id="e684c-353">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="e684c-353">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e684c-354">要件</span><span class="sxs-lookup"><span data-stu-id="e684c-354">Requirements</span></span>

|<span data-ttu-id="e684c-355">要件</span><span class="sxs-lookup"><span data-stu-id="e684c-355">Requirement</span></span>| <span data-ttu-id="e684c-356">値</span><span class="sxs-lookup"><span data-stu-id="e684c-356">Value</span></span>|
|---|---|
|[<span data-ttu-id="e684c-357">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e684c-357">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e684c-358">1.3</span><span class="sxs-lookup"><span data-stu-id="e684c-358">1.3</span></span>|
|[<span data-ttu-id="e684c-359">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e684c-359">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e684c-360">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e684c-360">ReadItem</span></span>|
|[<span data-ttu-id="e684c-361">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e684c-361">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e684c-362">新規作成と閲覧</span><span class="sxs-lookup"><span data-stu-id="e684c-362">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="e684c-363">例</span><span class="sxs-lookup"><span data-stu-id="e684c-363">Example</span></span>

```javascript
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="e684c-364">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="e684c-364">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="e684c-365">ユーザーと Office アドインを識別するトークンを取得します。</span><span class="sxs-lookup"><span data-stu-id="e684c-365">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="e684c-366">`getUserIdentityTokenAsync` メソッドは、[アドインとユーザーをサード パーティのシステムで識別して認証](https://docs.microsoft.com/outlook/add-ins/authentication)することのできるトークンを返します。</span><span class="sxs-lookup"><span data-stu-id="e684c-366">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](https://docs.microsoft.com/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="e684c-367">パラメーター</span><span class="sxs-lookup"><span data-stu-id="e684c-367">Parameters</span></span>

|<span data-ttu-id="e684c-368">名前</span><span class="sxs-lookup"><span data-stu-id="e684c-368">Name</span></span>| <span data-ttu-id="e684c-369">型</span><span class="sxs-lookup"><span data-stu-id="e684c-369">Type</span></span>| <span data-ttu-id="e684c-370">属性</span><span class="sxs-lookup"><span data-stu-id="e684c-370">Attributes</span></span>| <span data-ttu-id="e684c-371">説明</span><span class="sxs-lookup"><span data-stu-id="e684c-371">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="e684c-372">function</span><span class="sxs-lookup"><span data-stu-id="e684c-372">function</span></span>||<span data-ttu-id="e684c-373">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="e684c-373">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="e684c-374">トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="e684c-374">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="e684c-375">Object</span><span class="sxs-lookup"><span data-stu-id="e684c-375">Object</span></span>| <span data-ttu-id="e684c-376">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="e684c-376">&lt;optional&gt;</span></span>|<span data-ttu-id="e684c-377">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="e684c-377">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e684c-378">要件</span><span class="sxs-lookup"><span data-stu-id="e684c-378">Requirements</span></span>

|<span data-ttu-id="e684c-379">要件</span><span class="sxs-lookup"><span data-stu-id="e684c-379">Requirement</span></span>| <span data-ttu-id="e684c-380">値</span><span class="sxs-lookup"><span data-stu-id="e684c-380">Value</span></span>|
|---|---|
|[<span data-ttu-id="e684c-381">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e684c-381">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e684c-382">1.0</span><span class="sxs-lookup"><span data-stu-id="e684c-382">1.0</span></span>|
|[<span data-ttu-id="e684c-383">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e684c-383">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e684c-384">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e684c-384">ReadItem</span></span>|
|[<span data-ttu-id="e684c-385">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e684c-385">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e684c-386">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="e684c-386">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e684c-387">例</span><span class="sxs-lookup"><span data-stu-id="e684c-387">Example</span></span>

```javascript
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="e684c-388">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="e684c-388">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="e684c-389">ユーザーのメールボックスをホストしている Exchange サーバー上の Exchange Web サービス (EWS) のサービスに対して非同期の要求を行います。</span><span class="sxs-lookup"><span data-stu-id="e684c-389">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="e684c-390">このメソッドは、次のシナリオではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="e684c-390">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="e684c-391">Outlook for iOS または Outlook for Android を使用している場合</span><span class="sxs-lookup"><span data-stu-id="e684c-391">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="e684c-392">アドインが Gmail のメールボックスに読み込まれる場合</span><span class="sxs-lookup"><span data-stu-id="e684c-392">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="e684c-393">このような場合は、アドインでは [REST API を使用](https://docs.microsoft.com/outlook/add-ins/use-rest-api)して、代わりにユーザーのメールボックスにアクセスする必要があります。</span><span class="sxs-lookup"><span data-stu-id="e684c-393">In these cases, add-ins should [use REST APIs](https://docs.microsoft.com/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="e684c-394">`makeEwsRequestAsync` メソッドは、アドインの代わりに Exchange に EWS 要求を送信します。</span><span class="sxs-lookup"><span data-stu-id="e684c-394">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="e684c-395">サポートされている EWS 操作の一覧については、「[Outlook アドインからの Web サービスの呼び出し](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e684c-395">See [Call web services from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="e684c-396">`makeEwsRequestAsync` メソッドでは、フォルダー関連アイテムを要求できません。</span><span class="sxs-lookup"><span data-stu-id="e684c-396">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="e684c-397">XML 要求では UTF-8 エンコードを指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="e684c-397">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="e684c-p124">`makeEwsRequestAsync` メソッドを使用するには、アドインに **ReadWriteMailbox** アクセス許可が必要です。**ReadWriteMailbox** アクセス許可と、`makeEwsRequestAsync` メソッドで呼び出せる EWS 操作の使い方については、「[ユーザーのメールボックスへのメール アドイン アクセスのアクセス許可を指定する](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e684c-p124">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="e684c-400">サーバー管理者は、クライアント アクセス サーバーの EWS ディレクトリで `OAuthAuthentication` を true に設定して、`makeEwsRequestAsync` メソッドで EWS 要求を行うことができるようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="e684c-400">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="e684c-401">バージョンの相違点</span><span class="sxs-lookup"><span data-stu-id="e684c-401">Version differences</span></span>

<span data-ttu-id="e684c-402">バージョン 15.0.4535.1004 より前のバージョンの Outlook で実行しているメール アプリで `makeEwsRequestAsync` メソッドを使う場合は、エンコード値を `ISO-8859-1` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="e684c-402">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="e684c-p125">Outlook on the web でメール アプリを実行している場合は、エンコード値を設定する必要はありません。mailbox.diagnostics.hostName プロパティを使って、メール アプリを Outlook で実行しているのか、Outlook on the web で実行しているのかを確認できます。mailbox.diagnostics.hostVersion プロパティを使って、どのバージョンの Outlook を使って実行しているかを確認できます。</span><span class="sxs-lookup"><span data-stu-id="e684c-p125">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e684c-406">パラメーター</span><span class="sxs-lookup"><span data-stu-id="e684c-406">Parameters</span></span>

|<span data-ttu-id="e684c-407">名前</span><span class="sxs-lookup"><span data-stu-id="e684c-407">Name</span></span>| <span data-ttu-id="e684c-408">型</span><span class="sxs-lookup"><span data-stu-id="e684c-408">Type</span></span>| <span data-ttu-id="e684c-409">属性</span><span class="sxs-lookup"><span data-stu-id="e684c-409">Attributes</span></span>| <span data-ttu-id="e684c-410">説明</span><span class="sxs-lookup"><span data-stu-id="e684c-410">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="e684c-411">String</span><span class="sxs-lookup"><span data-stu-id="e684c-411">String</span></span>||<span data-ttu-id="e684c-412">EWS 要求です。</span><span class="sxs-lookup"><span data-stu-id="e684c-412">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="e684c-413">function</span><span class="sxs-lookup"><span data-stu-id="e684c-413">function</span></span>||<span data-ttu-id="e684c-414">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="e684c-414">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="e684c-415">EWS 呼び出しの XML 結果は、`asyncResult.value` プロパティ内の文字列として提供されています。</span><span class="sxs-lookup"><span data-stu-id="e684c-415">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="e684c-416">結果のサイズが 1 MB を超える場合、代わりにエラー メッセージが返されます。</span><span class="sxs-lookup"><span data-stu-id="e684c-416">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="e684c-417">Object</span><span class="sxs-lookup"><span data-stu-id="e684c-417">Object</span></span>| <span data-ttu-id="e684c-418">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="e684c-418">&lt;optional&gt;</span></span>|<span data-ttu-id="e684c-419">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="e684c-419">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e684c-420">要件</span><span class="sxs-lookup"><span data-stu-id="e684c-420">Requirements</span></span>

|<span data-ttu-id="e684c-421">要件</span><span class="sxs-lookup"><span data-stu-id="e684c-421">Requirement</span></span>| <span data-ttu-id="e684c-422">値</span><span class="sxs-lookup"><span data-stu-id="e684c-422">Value</span></span>|
|---|---|
|[<span data-ttu-id="e684c-423">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e684c-423">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e684c-424">1.0</span><span class="sxs-lookup"><span data-stu-id="e684c-424">1.0</span></span>|
|[<span data-ttu-id="e684c-425">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e684c-425">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e684c-426">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="e684c-426">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="e684c-427">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e684c-427">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e684c-428">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="e684c-428">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e684c-429">例</span><span class="sxs-lookup"><span data-stu-id="e684c-429">Example</span></span>

<span data-ttu-id="e684c-430">次の例は、`GetItem` 操作を使ってアイテムの件名を取得するため、`makeEwsRequestAsync` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="e684c-430">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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
