---
title: Office. メールボックス要件セット1.3
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 18e144702689c1df2b89f53c666932a74bc795ef
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/27/2019
ms.locfileid: "30872054"
---
# <a name="mailbox"></a><span data-ttu-id="f842a-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="f842a-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="f842a-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="f842a-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="f842a-104">Microsoft Outlook と Microsoft Outlook on the web の Outlook アドイン オブジェクト モデルへのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="f842a-104">Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f842a-105">要件</span><span class="sxs-lookup"><span data-stu-id="f842a-105">Requirements</span></span>

|<span data-ttu-id="f842a-106">要件</span><span class="sxs-lookup"><span data-stu-id="f842a-106">Requirement</span></span>| <span data-ttu-id="f842a-107">値</span><span class="sxs-lookup"><span data-stu-id="f842a-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="f842a-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f842a-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f842a-109">1.0</span><span class="sxs-lookup"><span data-stu-id="f842a-109">1.0</span></span>|
|[<span data-ttu-id="f842a-110">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f842a-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f842a-111">制限あり</span><span class="sxs-lookup"><span data-stu-id="f842a-111">Restricted</span></span>|
|[<span data-ttu-id="f842a-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f842a-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f842a-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f842a-113">Compose or Read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="f842a-114">名前空間</span><span class="sxs-lookup"><span data-stu-id="f842a-114">Namespaces</span></span>

<span data-ttu-id="f842a-115">[diagnostics](Office.context.mailbox.diagnostics.md): Outlook アドインに診断情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="f842a-115">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="f842a-116">[item](Office.context.mailbox.item.md):Outlook アドインのメッセージや予定にアクセスするためのメソッドとプロパティを提供します。</span><span class="sxs-lookup"><span data-stu-id="f842a-116">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="f842a-117">[userProfile](Office.context.mailbox.userProfile.md):Outlook アドインのユーザーに関する情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="f842a-117">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="f842a-118">メンバー</span><span class="sxs-lookup"><span data-stu-id="f842a-118">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="f842a-119">ewsUrl: 文字列</span><span class="sxs-lookup"><span data-stu-id="f842a-119">ewsUrl :String</span></span>

<span data-ttu-id="f842a-p101">このメール アカウントの Exchange Web サービス (EWS) エンドポイントの URL を取得します。読み取りモードのみです。</span><span class="sxs-lookup"><span data-stu-id="f842a-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="f842a-122">このメンバーは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f842a-122">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="f842a-p102">`ewsUrl` 値は、リモート サービスで、ユーザーのメールボックスに EWS 呼び出しを行うために使用することができます。たとえば、[選択した項目から添付ファイルを取得する](/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="f842a-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="f842a-125">閲覧モードで `ewsUrl` メンバーを呼び出すには、アプリでマニフェスト内に **ReadItem** アクセス許可を指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="f842a-125">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="f842a-p103">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback)メソッドを呼び出してから、`ewsUrl`メンバーを使用する必要があります。アプリには、`saveAsync`メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="f842a-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="f842a-128">型</span><span class="sxs-lookup"><span data-stu-id="f842a-128">Type</span></span>

*   <span data-ttu-id="f842a-129">String</span><span class="sxs-lookup"><span data-stu-id="f842a-129">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f842a-130">要件</span><span class="sxs-lookup"><span data-stu-id="f842a-130">Requirements</span></span>

|<span data-ttu-id="f842a-131">要件</span><span class="sxs-lookup"><span data-stu-id="f842a-131">Requirement</span></span>| <span data-ttu-id="f842a-132">値</span><span class="sxs-lookup"><span data-stu-id="f842a-132">Value</span></span>|
|---|---|
|[<span data-ttu-id="f842a-133">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f842a-133">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f842a-134">1.0</span><span class="sxs-lookup"><span data-stu-id="f842a-134">1.0</span></span>|
|[<span data-ttu-id="f842a-135">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f842a-135">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f842a-136">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f842a-136">ReadItem</span></span>|
|[<span data-ttu-id="f842a-137">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f842a-137">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f842a-138">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f842a-138">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="f842a-139">メソッド</span><span class="sxs-lookup"><span data-stu-id="f842a-139">Methods</span></span>

####  <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="f842a-140">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="f842a-140">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="f842a-141">REST 形式のアイテム ID を EWS 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="f842a-141">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="f842a-142">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f842a-142">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="f842a-p104">REST API ([Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) や [Microsoft Graph](https://graph.microsoft.io/) など) 経由で取得された項目 ID は、Exchange Web サービス (EWS) で使用される形式とは異なる形式を使用します。`convertToEwsId` メソッドは、REST 形式の ID を EWS 用の適切な形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="f842a-p104">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f842a-145">パラメーター</span><span class="sxs-lookup"><span data-stu-id="f842a-145">Parameters</span></span>

|<span data-ttu-id="f842a-146">名前</span><span class="sxs-lookup"><span data-stu-id="f842a-146">Name</span></span>| <span data-ttu-id="f842a-147">種類</span><span class="sxs-lookup"><span data-stu-id="f842a-147">Type</span></span>| <span data-ttu-id="f842a-148">説明</span><span class="sxs-lookup"><span data-stu-id="f842a-148">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="f842a-149">String</span><span class="sxs-lookup"><span data-stu-id="f842a-149">String</span></span>|<span data-ttu-id="f842a-150">Outlook REST API 用に書式設定された項目 ID</span><span class="sxs-lookup"><span data-stu-id="f842a-150">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="f842a-151">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="f842a-151">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_3/office.mailboxenums.restversion)|<span data-ttu-id="f842a-152">項目 ID の取得に使用された Outlook REST API のバージョンを示す値。</span><span class="sxs-lookup"><span data-stu-id="f842a-152">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f842a-153">要件</span><span class="sxs-lookup"><span data-stu-id="f842a-153">Requirements</span></span>

|<span data-ttu-id="f842a-154">要件</span><span class="sxs-lookup"><span data-stu-id="f842a-154">Requirement</span></span>| <span data-ttu-id="f842a-155">値</span><span class="sxs-lookup"><span data-stu-id="f842a-155">Value</span></span>|
|---|---|
|[<span data-ttu-id="f842a-156">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f842a-156">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f842a-157">1.3</span><span class="sxs-lookup"><span data-stu-id="f842a-157">1.3</span></span>|
|[<span data-ttu-id="f842a-158">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f842a-158">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f842a-159">制限あり</span><span class="sxs-lookup"><span data-stu-id="f842a-159">Restricted</span></span>|
|[<span data-ttu-id="f842a-160">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f842a-160">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f842a-161">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f842a-161">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f842a-162">戻り値:</span><span class="sxs-lookup"><span data-stu-id="f842a-162">Returns:</span></span>

<span data-ttu-id="f842a-163">種類: 文字列</span><span class="sxs-lookup"><span data-stu-id="f842a-163">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="f842a-164">例</span><span class="sxs-lookup"><span data-stu-id="f842a-164">Example</span></span>

```javascript
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook13officelocalclienttime"></a><span data-ttu-id="f842a-165">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_3/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="f842a-165">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_3/office.LocalClientTime)}</span></span>

<span data-ttu-id="f842a-166">クライアントの現地時間の時間情報が含まれている辞書を取得します。</span><span class="sxs-lookup"><span data-stu-id="f842a-166">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="f842a-p105">Outlook 用メール アプリや Outlook Web App で使う日付と時刻は、異なるタイム ゾーンを使うことができます。Outlook では、クライアント コンピューターのタイム ゾーンを使います。Outlook Web App では、Exchange 管理センター (EAC) で設定されたタイム ゾーンを使います。ユーザー インターフェイスに表示される値が、常にユーザーが期待するタイム ゾーンと一致するように日付と時刻の値を処理する必要があります。</span><span class="sxs-lookup"><span data-stu-id="f842a-p105">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="f842a-p106">Outlook でメール アプリが実行されている場合、`convertToLocalClientTime` メソッドは、クライアント コンピューターのタイム ゾーンに設定された値で辞書オブジェクトを返します。Outlook Web Apps でメール アプリが実行されている場合、`convertToLocalClientTime` メソッドは、EAC で指定されたタイム ゾーンに設定された値で辞書オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="f842a-p106">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f842a-172">パラメーター</span><span class="sxs-lookup"><span data-stu-id="f842a-172">Parameters</span></span>

|<span data-ttu-id="f842a-173">名前</span><span class="sxs-lookup"><span data-stu-id="f842a-173">Name</span></span>| <span data-ttu-id="f842a-174">種類</span><span class="sxs-lookup"><span data-stu-id="f842a-174">Type</span></span>| <span data-ttu-id="f842a-175">説明</span><span class="sxs-lookup"><span data-stu-id="f842a-175">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="f842a-176">日付</span><span class="sxs-lookup"><span data-stu-id="f842a-176">Date</span></span>|<span data-ttu-id="f842a-177">Date オブジェクト</span><span class="sxs-lookup"><span data-stu-id="f842a-177">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f842a-178">要件</span><span class="sxs-lookup"><span data-stu-id="f842a-178">Requirements</span></span>

|<span data-ttu-id="f842a-179">要件</span><span class="sxs-lookup"><span data-stu-id="f842a-179">Requirement</span></span>| <span data-ttu-id="f842a-180">値</span><span class="sxs-lookup"><span data-stu-id="f842a-180">Value</span></span>|
|---|---|
|[<span data-ttu-id="f842a-181">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f842a-181">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f842a-182">1.0</span><span class="sxs-lookup"><span data-stu-id="f842a-182">1.0</span></span>|
|[<span data-ttu-id="f842a-183">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f842a-183">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f842a-184">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f842a-184">ReadItem</span></span>|
|[<span data-ttu-id="f842a-185">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f842a-185">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f842a-186">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f842a-186">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f842a-187">戻り値 :</span><span class="sxs-lookup"><span data-stu-id="f842a-187">Returns:</span></span>

<span data-ttu-id="f842a-188">種類:[LocalClientTime](/javascript/api/outlook_1_3/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="f842a-188">Type: [LocalClientTime](/javascript/api/outlook_1_3/office.LocalClientTime)</span></span>

####  <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="f842a-189">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="f842a-189">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="f842a-190">EWS 形式のアイテム ID を REST 形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="f842a-190">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="f842a-191">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f842a-191">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="f842a-p107">EWS 経由または `itemId` プロパティ経由で取得される項目 ID では、REST API ([Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) または [Microsoft Graph](https://graph.microsoft.io/) など) で使用される形式とは異なる形式を使用します。`convertToRestId` メソッドは、EWS 形式の ID を REST 用の適切な形式に変換します。</span><span class="sxs-lookup"><span data-stu-id="f842a-p107">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f842a-194">パラメーター</span><span class="sxs-lookup"><span data-stu-id="f842a-194">Parameters</span></span>

|<span data-ttu-id="f842a-195">名前</span><span class="sxs-lookup"><span data-stu-id="f842a-195">Name</span></span>| <span data-ttu-id="f842a-196">種類</span><span class="sxs-lookup"><span data-stu-id="f842a-196">Type</span></span>| <span data-ttu-id="f842a-197">説明</span><span class="sxs-lookup"><span data-stu-id="f842a-197">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="f842a-198">文字列</span><span class="sxs-lookup"><span data-stu-id="f842a-198">String</span></span>|<span data-ttu-id="f842a-199">Exchange Web サービス (EWS) 用に書式設定された項目 ID</span><span class="sxs-lookup"><span data-stu-id="f842a-199">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="f842a-200">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="f842a-200">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_3/office.mailboxenums.restversion)|<span data-ttu-id="f842a-201">変換後の ID とともに使用される Outlook REST API のバージョンを示す値。</span><span class="sxs-lookup"><span data-stu-id="f842a-201">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f842a-202">要件</span><span class="sxs-lookup"><span data-stu-id="f842a-202">Requirements</span></span>

|<span data-ttu-id="f842a-203">要件</span><span class="sxs-lookup"><span data-stu-id="f842a-203">Requirement</span></span>| <span data-ttu-id="f842a-204">値</span><span class="sxs-lookup"><span data-stu-id="f842a-204">Value</span></span>|
|---|---|
|[<span data-ttu-id="f842a-205">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f842a-205">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f842a-206">1.3</span><span class="sxs-lookup"><span data-stu-id="f842a-206">1.3</span></span>|
|[<span data-ttu-id="f842a-207">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f842a-207">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f842a-208">制限あり</span><span class="sxs-lookup"><span data-stu-id="f842a-208">Restricted</span></span>|
|[<span data-ttu-id="f842a-209">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f842a-209">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f842a-210">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f842a-210">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f842a-211">戻り値:</span><span class="sxs-lookup"><span data-stu-id="f842a-211">Returns:</span></span>

<span data-ttu-id="f842a-212">種類: 文字列</span><span class="sxs-lookup"><span data-stu-id="f842a-212">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="f842a-213">例</span><span class="sxs-lookup"><span data-stu-id="f842a-213">Example</span></span>

```javascript
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="f842a-214">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="f842a-214">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="f842a-215">時間情報が含まれている辞書から Date オブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="f842a-215">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="f842a-216">`convertToUtcClientTime` メソッドは、ローカルの日付と時刻が含まれる辞書を、ローカルの日付と時刻の正しい値をもつ Date オブジェクトに変換します。</span><span class="sxs-lookup"><span data-stu-id="f842a-216">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f842a-217">パラメーター</span><span class="sxs-lookup"><span data-stu-id="f842a-217">Parameters</span></span>

|<span data-ttu-id="f842a-218">名前</span><span class="sxs-lookup"><span data-stu-id="f842a-218">Name</span></span>| <span data-ttu-id="f842a-219">型</span><span class="sxs-lookup"><span data-stu-id="f842a-219">Type</span></span>| <span data-ttu-id="f842a-220">説明</span><span class="sxs-lookup"><span data-stu-id="f842a-220">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="f842a-221">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="f842a-221">LocalClientTime</span></span>](/javascript/api/outlook_1_3/office.LocalClientTime)|<span data-ttu-id="f842a-222">変換するローカル時刻の値。</span><span class="sxs-lookup"><span data-stu-id="f842a-222">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f842a-223">要件</span><span class="sxs-lookup"><span data-stu-id="f842a-223">Requirements</span></span>

|<span data-ttu-id="f842a-224">要件</span><span class="sxs-lookup"><span data-stu-id="f842a-224">Requirement</span></span>| <span data-ttu-id="f842a-225">値</span><span class="sxs-lookup"><span data-stu-id="f842a-225">Value</span></span>|
|---|---|
|[<span data-ttu-id="f842a-226">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f842a-226">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f842a-227">1.0</span><span class="sxs-lookup"><span data-stu-id="f842a-227">1.0</span></span>|
|[<span data-ttu-id="f842a-228">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f842a-228">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f842a-229">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f842a-229">ReadItem</span></span>|
|[<span data-ttu-id="f842a-230">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f842a-230">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f842a-231">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f842a-231">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f842a-232">戻り値 :</span><span class="sxs-lookup"><span data-stu-id="f842a-232">Returns:</span></span>

<span data-ttu-id="f842a-233">時間が UTC で表現された Date オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="f842a-233">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="f842a-234">

<dt>種類</dt>

</span><span class="sxs-lookup"><span data-stu-id="f842a-234">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="f842a-235">Date</span><span class="sxs-lookup"><span data-stu-id="f842a-235">Date</span></span></dd>

</dl>

####  <a name="displayappointmentformitemid"></a><span data-ttu-id="f842a-236">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="f842a-236">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="f842a-237">既存の予定を表示します。</span><span class="sxs-lookup"><span data-stu-id="f842a-237">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="f842a-238">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f842a-238">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="f842a-239">`displayAppointmentForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存の予定を開きます。</span><span class="sxs-lookup"><span data-stu-id="f842a-239">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="f842a-p108">Outlook for Mac では、このメソッドを使用して、定期的に繰り返される予定の一部ではない単発の予定、または定期的に繰り替えされる予定の元の予定を表示できます。ただし、一連の予定のインスタンスは表示できません。これは、Outlook for Mac では、定期的に繰り返されるインスタンスのプロパティ  (項目 ID を含む) にアクセスできないためです。</span><span class="sxs-lookup"><span data-stu-id="f842a-p108">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="f842a-242">Outlook Web App では、このメソッドで、指定されたフォームの本文が 32 KB 以下の文字数の場合に、そのフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="f842a-242">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="f842a-243">指定した項目識別子が既存の予定を識別しない場合には、クライアント コンピュータまたはデバイスで空白のウィンドウが開き、エラー メッセージは返されません。</span><span class="sxs-lookup"><span data-stu-id="f842a-243">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f842a-244">パラメーター</span><span class="sxs-lookup"><span data-stu-id="f842a-244">Parameters</span></span>

|<span data-ttu-id="f842a-245">名前</span><span class="sxs-lookup"><span data-stu-id="f842a-245">Name</span></span>| <span data-ttu-id="f842a-246">型</span><span class="sxs-lookup"><span data-stu-id="f842a-246">Type</span></span>| <span data-ttu-id="f842a-247">説明</span><span class="sxs-lookup"><span data-stu-id="f842a-247">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="f842a-248">String</span><span class="sxs-lookup"><span data-stu-id="f842a-248">String</span></span>|<span data-ttu-id="f842a-249">既存の予定表の予定の Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="f842a-249">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f842a-250">要件</span><span class="sxs-lookup"><span data-stu-id="f842a-250">Requirements</span></span>

|<span data-ttu-id="f842a-251">要件</span><span class="sxs-lookup"><span data-stu-id="f842a-251">Requirement</span></span>| <span data-ttu-id="f842a-252">値</span><span class="sxs-lookup"><span data-stu-id="f842a-252">Value</span></span>|
|---|---|
|[<span data-ttu-id="f842a-253">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f842a-253">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f842a-254">1.0</span><span class="sxs-lookup"><span data-stu-id="f842a-254">1.0</span></span>|
|[<span data-ttu-id="f842a-255">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f842a-255">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f842a-256">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f842a-256">ReadItem</span></span>|
|[<span data-ttu-id="f842a-257">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f842a-257">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f842a-258">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f842a-258">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f842a-259">例</span><span class="sxs-lookup"><span data-stu-id="f842a-259">Example</span></span>

```javascript
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  <a name="displaymessageformitemid"></a><span data-ttu-id="f842a-260">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="f842a-260">displayMessageForm(itemId)</span></span>

<span data-ttu-id="f842a-261">既存のメッセージを表示します。</span><span class="sxs-lookup"><span data-stu-id="f842a-261">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="f842a-262">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f842a-262">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="f842a-263">`displayMessageForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存のメッセージを開きます。</span><span class="sxs-lookup"><span data-stu-id="f842a-263">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="f842a-264">Outlook Web App では、このメソッドは、指定されたフォームの本文が 32 KB 以下の文字数の場合に、そのフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="f842a-264">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="f842a-265">指定した項目識別子が既存のメッセージを識別しない場合には、クラアント コンピュータでメッセージは表示されず、エラー メッセージは返されません。</span><span class="sxs-lookup"><span data-stu-id="f842a-265">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="f842a-p109">予定を表す `itemId` が含まれる `displayMessageForm` を使用しないでください。`displayAppointmentForm` メソッドを使用して既存の予定を表示し、`displayNewAppointmentForm` を使用して新しい予定を作成するフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="f842a-p109">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f842a-268">パラメーター</span><span class="sxs-lookup"><span data-stu-id="f842a-268">Parameters</span></span>

|<span data-ttu-id="f842a-269">名前</span><span class="sxs-lookup"><span data-stu-id="f842a-269">Name</span></span>| <span data-ttu-id="f842a-270">型</span><span class="sxs-lookup"><span data-stu-id="f842a-270">Type</span></span>| <span data-ttu-id="f842a-271">説明</span><span class="sxs-lookup"><span data-stu-id="f842a-271">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="f842a-272">String</span><span class="sxs-lookup"><span data-stu-id="f842a-272">String</span></span>|<span data-ttu-id="f842a-273">既存のメッセージの Exchange Web サービス(EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="f842a-273">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f842a-274">要件</span><span class="sxs-lookup"><span data-stu-id="f842a-274">Requirements</span></span>

|<span data-ttu-id="f842a-275">要件</span><span class="sxs-lookup"><span data-stu-id="f842a-275">Requirement</span></span>| <span data-ttu-id="f842a-276">値</span><span class="sxs-lookup"><span data-stu-id="f842a-276">Value</span></span>|
|---|---|
|[<span data-ttu-id="f842a-277">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f842a-277">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f842a-278">1.0</span><span class="sxs-lookup"><span data-stu-id="f842a-278">1.0</span></span>|
|[<span data-ttu-id="f842a-279">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f842a-279">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f842a-280">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f842a-280">ReadItem</span></span>|
|[<span data-ttu-id="f842a-281">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f842a-281">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f842a-282">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f842a-282">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f842a-283">例</span><span class="sxs-lookup"><span data-stu-id="f842a-283">Example</span></span>

```javascript
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="f842a-284">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="f842a-284">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="f842a-285">新しい予定を作成するためのフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="f842a-285">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="f842a-286">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f842a-286">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="f842a-p110">`displayNewAppointmentForm` メソッドを使用すると、ユーザーが新しい予定または会議を作成できるフォームが開きます。パラメータを指定すると、予定のフォーム フィールドにパラメータの内容が自動的に入力されます。</span><span class="sxs-lookup"><span data-stu-id="f842a-p110">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="f842a-p111">Outlook Web App および OWA for Devices では、このメソッドは出席者フィールドが含まれるフォームを常に表示します。入力因数として出席者を指定しない場合には、このメソッドは [**保存**] ボタンのあるフォームを表示します。出席者を指定した場合には、フォームにはその出席者と [**送信**] ボタンが含まれます。</span><span class="sxs-lookup"><span data-stu-id="f842a-p111">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="f842a-p112">Outlook リッチ クライアントおよび Outlook RT では、`requiredAttendees`、`optionalAttendees` または `resources` パラメータに出席者またはリソースを指定した場合、このメソッドは [**送信**] ボタンがある会議フォームを表示します。受信者を指定しない場合には、このメソッドは [**保存して閉じる**] ボタンがある予定フォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="f842a-p112">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="f842a-294">パラメータのいずれかが指定されたサイズ制限を超えた場合、または不明なパラメータ名が指定された場合には、例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="f842a-294">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f842a-295">パラメーター</span><span class="sxs-lookup"><span data-stu-id="f842a-295">Parameters</span></span>

|<span data-ttu-id="f842a-296">名前</span><span class="sxs-lookup"><span data-stu-id="f842a-296">Name</span></span>| <span data-ttu-id="f842a-297">型</span><span class="sxs-lookup"><span data-stu-id="f842a-297">Type</span></span>| <span data-ttu-id="f842a-298">説明</span><span class="sxs-lookup"><span data-stu-id="f842a-298">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="f842a-299">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="f842a-299">Object</span></span> | <span data-ttu-id="f842a-300">新しい予定を記述するパラメータの辞書。</span><span class="sxs-lookup"><span data-stu-id="f842a-300">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="f842a-301">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="f842a-301">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="f842a-p113">予定への各必須出席者の電子メール アドレスを含む文字列の配列、または出席者の `EmailAddressDetails` オブジェクトを含む配列。この配列は、最大 100 エントリまでに制限されています。</span><span class="sxs-lookup"><span data-stu-id="f842a-p113">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="f842a-304">配列。&lt;文字列&gt; | 配列。&lt;[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="f842a-304">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="f842a-p114">予定への各任意出席者の電子メール アドレスを含む文字列の配列、または出席者の `EmailAddressDetails` オブジェクトを含む配列。この配列は、最大 100 エントリまでに制限されています。</span><span class="sxs-lookup"><span data-stu-id="f842a-p114">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="f842a-307">日付</span><span class="sxs-lookup"><span data-stu-id="f842a-307">Date</span></span> | <span data-ttu-id="f842a-308">予定の開始日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="f842a-308">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="f842a-309">日付</span><span class="sxs-lookup"><span data-stu-id="f842a-309">Date</span></span> | <span data-ttu-id="f842a-310">予定の終了日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="f842a-310">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="f842a-311">String</span><span class="sxs-lookup"><span data-stu-id="f842a-311">String</span></span> | <span data-ttu-id="f842a-p115">予定の場所を含む文字列。文字列は最大 255 文字までに制限されています。</span><span class="sxs-lookup"><span data-stu-id="f842a-p115">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="f842a-314">配列。&lt; 文字列&gt;</span><span class="sxs-lookup"><span data-stu-id="f842a-314">Array.&lt;String&gt;</span></span> | <span data-ttu-id="f842a-p116">予定に必要なリソースを含む文字列の配列。配列は最大 100 エントリまでに制限されています。</span><span class="sxs-lookup"><span data-stu-id="f842a-p116">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="f842a-317">String</span><span class="sxs-lookup"><span data-stu-id="f842a-317">String</span></span> | <span data-ttu-id="f842a-p117">予定の件名を含む文字列。文字列は最大 255 文字までに制限されています。</span><span class="sxs-lookup"><span data-stu-id="f842a-p117">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="f842a-320">文字列</span><span class="sxs-lookup"><span data-stu-id="f842a-320">String</span></span> | <span data-ttu-id="f842a-p118">予定の本文。本文の内容は、最大サイズが 32 KB までに制限されています。</span><span class="sxs-lookup"><span data-stu-id="f842a-p118">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f842a-323">要件</span><span class="sxs-lookup"><span data-stu-id="f842a-323">Requirements</span></span>

|<span data-ttu-id="f842a-324">要件</span><span class="sxs-lookup"><span data-stu-id="f842a-324">Requirement</span></span>| <span data-ttu-id="f842a-325">値</span><span class="sxs-lookup"><span data-stu-id="f842a-325">Value</span></span>|
|---|---|
|[<span data-ttu-id="f842a-326">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f842a-326">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f842a-327">1.0</span><span class="sxs-lookup"><span data-stu-id="f842a-327">1.0</span></span>|
|[<span data-ttu-id="f842a-328">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f842a-328">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f842a-329">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f842a-329">ReadItem</span></span>|
|[<span data-ttu-id="f842a-330">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f842a-330">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f842a-331">読み取り</span><span class="sxs-lookup"><span data-stu-id="f842a-331">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f842a-332">例</span><span class="sxs-lookup"><span data-stu-id="f842a-332">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="f842a-333">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="f842a-333">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="f842a-334">Exchange Server から添付ファイルやアイテムを取得するために使用するトークンを含む文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="f842a-334">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="f842a-p119">`getCallbackTokenAsync` メソッドは、非同期の呼び出しを行なって、ユーザーのメールボックスをホストする Exchange Server から opaque トークンを取得します。コールバック トークンの有効期間は 5 分です。</span><span class="sxs-lookup"><span data-stu-id="f842a-p119">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="f842a-p120">このトークンと、添付ファイル識別子または項目識別子は、サードパーティーのシステムに渡すことができます。サードパーティーのシステムでは、添付ファイルまたは項目を返すための Exchange Web サービス (EWS) の [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) または [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) 操作を呼び出すのに、トークンをベアラー承認トークンとして使用します。たとえば、リモート サービスを作成して[選択した項目から添付ファイルを取得](/outlook/add-ins/get-attachments-of-an-outlook-item)することができます。</span><span class="sxs-lookup"><span data-stu-id="f842a-p120">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="f842a-340">アプリでは、閲覧モードで `getCallbackTokenAsync` メソッドを呼び出すために、 **ReadItem** アクセス許可をアプリのマニフェストで指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="f842a-340">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="f842a-p121">新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出して、`getCallbackTokenAsync` メソッドに渡すための項目識別子を取得する必要があります。アプリには、`saveAsync` メソッドを呼び出すために **ReadWriteItem** アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="f842a-p121">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f842a-343">パラメーター</span><span class="sxs-lookup"><span data-stu-id="f842a-343">Parameters</span></span>

|<span data-ttu-id="f842a-344">名前</span><span class="sxs-lookup"><span data-stu-id="f842a-344">Name</span></span>| <span data-ttu-id="f842a-345">型</span><span class="sxs-lookup"><span data-stu-id="f842a-345">Type</span></span>| <span data-ttu-id="f842a-346">属性</span><span class="sxs-lookup"><span data-stu-id="f842a-346">Attributes</span></span>| <span data-ttu-id="f842a-347">説明</span><span class="sxs-lookup"><span data-stu-id="f842a-347">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="f842a-348">function</span><span class="sxs-lookup"><span data-stu-id="f842a-348">function</span></span>||<span data-ttu-id="f842a-p122">メソッドが完了すると、`callback` パラメータに渡された関数が、単一パラメータ `asyncResult` で呼び出されます。このパラメータは、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトです。トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="f842a-p122">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="f842a-351">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="f842a-351">Object</span></span>| <span data-ttu-id="f842a-352">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="f842a-352">&lt;optional&gt;</span></span>|<span data-ttu-id="f842a-353">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="f842a-353">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f842a-354">要件</span><span class="sxs-lookup"><span data-stu-id="f842a-354">Requirements</span></span>

|<span data-ttu-id="f842a-355">要件</span><span class="sxs-lookup"><span data-stu-id="f842a-355">Requirement</span></span>| <span data-ttu-id="f842a-356">値</span><span class="sxs-lookup"><span data-stu-id="f842a-356">Value</span></span>|
|---|---|
|[<span data-ttu-id="f842a-357">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f842a-357">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f842a-358">1.3</span><span class="sxs-lookup"><span data-stu-id="f842a-358">1.3</span></span>|
|[<span data-ttu-id="f842a-359">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f842a-359">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f842a-360">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f842a-360">ReadItem</span></span>|
|[<span data-ttu-id="f842a-361">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f842a-361">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f842a-362">新規作成と閲覧</span><span class="sxs-lookup"><span data-stu-id="f842a-362">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="f842a-363">例</span><span class="sxs-lookup"><span data-stu-id="f842a-363">Example</span></span>

```javascript
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="f842a-364">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="f842a-364">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="f842a-365">ユーザーと Office アドインを識別するトークンを取得します。</span><span class="sxs-lookup"><span data-stu-id="f842a-365">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="f842a-366">`getUserIdentityTokenAsync` メソッドは、[アドインとユーザーをサードパーティのシステムで識別して認証](/outlook/add-ins/authentication)するのに使用できるトークンを返します。</span><span class="sxs-lookup"><span data-stu-id="f842a-366">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="f842a-367">パラメーター</span><span class="sxs-lookup"><span data-stu-id="f842a-367">Parameters</span></span>

|<span data-ttu-id="f842a-368">名前</span><span class="sxs-lookup"><span data-stu-id="f842a-368">Name</span></span>| <span data-ttu-id="f842a-369">型</span><span class="sxs-lookup"><span data-stu-id="f842a-369">Type</span></span>| <span data-ttu-id="f842a-370">属性</span><span class="sxs-lookup"><span data-stu-id="f842a-370">Attributes</span></span>| <span data-ttu-id="f842a-371">説明</span><span class="sxs-lookup"><span data-stu-id="f842a-371">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="f842a-372">function</span><span class="sxs-lookup"><span data-stu-id="f842a-372">function</span></span>||<span data-ttu-id="f842a-373">メソッドが完了すると、`callback` パラメータに渡された関数が、単一パラメータ `asyncResult` で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="f842a-373">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="f842a-374">トークンは、`asyncResult.value`プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="f842a-374">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="f842a-375">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="f842a-375">Object</span></span>| <span data-ttu-id="f842a-376">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="f842a-376">&lt;optional&gt;</span></span>|<span data-ttu-id="f842a-377">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="f842a-377">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f842a-378">要件</span><span class="sxs-lookup"><span data-stu-id="f842a-378">Requirements</span></span>

|<span data-ttu-id="f842a-379">要件</span><span class="sxs-lookup"><span data-stu-id="f842a-379">Requirement</span></span>| <span data-ttu-id="f842a-380">値</span><span class="sxs-lookup"><span data-stu-id="f842a-380">Value</span></span>|
|---|---|
|[<span data-ttu-id="f842a-381">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f842a-381">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f842a-382">1.0</span><span class="sxs-lookup"><span data-stu-id="f842a-382">1.0</span></span>|
|[<span data-ttu-id="f842a-383">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f842a-383">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f842a-384">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f842a-384">ReadItem</span></span>|
|[<span data-ttu-id="f842a-385">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f842a-385">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f842a-386">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f842a-386">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f842a-387">例</span><span class="sxs-lookup"><span data-stu-id="f842a-387">Example</span></span>

```javascript
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="f842a-388">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="f842a-388">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="f842a-389">ユーザーのメールボックスをホストしている Exchange サーバー上の Exchange Web サービス (EWS) のサービスに対して非同期の要求を行います。</span><span class="sxs-lookup"><span data-stu-id="f842a-389">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="f842a-390">このメソッドは、次のシナリオではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f842a-390">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="f842a-391">Outlook for iOS または Outlook for Android で</span><span class="sxs-lookup"><span data-stu-id="f842a-391">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="f842a-392">アドインが Gmail のメールボックスにロードされる場合</span><span class="sxs-lookup"><span data-stu-id="f842a-392">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="f842a-393">これらの場合、アドインではユーザーのメールボックスにアクセスするために、代わりに [REST API を使用する](/outlook/add-ins/use-rest-api)必要があります。</span><span class="sxs-lookup"><span data-stu-id="f842a-393">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="f842a-p123">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange. See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span><span class="sxs-lookup"><span data-stu-id="f842a-p123">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange. See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="f842a-396">`makeEwsRequestAsync` メソッドでは、フォルダー関連アイテムを要求できません。</span><span class="sxs-lookup"><span data-stu-id="f842a-396">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="f842a-397">XML 要求では、UTF-8 エンコードを指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="f842a-397">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="f842a-p124">アドインには、`makeEwsRequestAsync` メソッドを使用するために **ReadWriteMailbox** アクセス許可が必要です。**ReadWriteMailbox** アクセス許可と、`makeEwsRequestAsync` メソッドで呼び出すことのできる EWS 操作の使用の詳細については、「[ユーザーのメールボックスへのメール アドイン アクセスのアクセス許可を指定する](/outlook/add-ins/understanding-outlook-add-in-permissions)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f842a-p124">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="f842a-400">サーバー管理者は、クライアント アクセス サーバーの EWS ディレクトリで `OAuthAuthentication` を true に設定して、`makeEwsRequestAsync` メソッドで EWS 要求を行うことができるようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="f842a-400">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="f842a-401">バージョンの相違点</span><span class="sxs-lookup"><span data-stu-id="f842a-401">Version differences</span></span>

<span data-ttu-id="f842a-402">バージョン 15.0.4535.1004 より前のバージョンの Outlook で実行しているメール アプリで `makeEwsRequestAsync` メソッドを使用する場合には、エンコード値を `ISO-8859-1` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="f842a-402">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="f842a-p125">メール アプリが Outlook on the web で実行されている場合には、エンコード値を設定する必要はありません。メールボックスを使用してメール アプリが Outlook で実行されているのか、Outlook on the web で実行されているのかを判断する必要があります。mailbox.diagnostics.hostVersion プロパティを使用すれば、どのバージョンの Outlook が実行されているのかがわかります。</span><span class="sxs-lookup"><span data-stu-id="f842a-p125">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f842a-406">パラメーター</span><span class="sxs-lookup"><span data-stu-id="f842a-406">Parameters</span></span>

|<span data-ttu-id="f842a-407">名前</span><span class="sxs-lookup"><span data-stu-id="f842a-407">Name</span></span>| <span data-ttu-id="f842a-408">型</span><span class="sxs-lookup"><span data-stu-id="f842a-408">Type</span></span>| <span data-ttu-id="f842a-409">属性</span><span class="sxs-lookup"><span data-stu-id="f842a-409">Attributes</span></span>| <span data-ttu-id="f842a-410">説明</span><span class="sxs-lookup"><span data-stu-id="f842a-410">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="f842a-411">文字列</span><span class="sxs-lookup"><span data-stu-id="f842a-411">String</span></span>||<span data-ttu-id="f842a-412">EWS 要求。</span><span class="sxs-lookup"><span data-stu-id="f842a-412">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="f842a-413">関数</span><span class="sxs-lookup"><span data-stu-id="f842a-413">function</span></span>||<span data-ttu-id="f842a-414">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="f842a-414">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="f842a-p126">The XML result of the EWS call is provided as a string in the `asyncResult.value` property. If the result exceeds 1 MB in size, an error message is returned instead.</span><span class="sxs-lookup"><span data-stu-id="f842a-p126">The XML result of the EWS call is provided as a string in the `asyncResult.value` property. If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="f842a-417">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="f842a-417">Object</span></span>| <span data-ttu-id="f842a-418">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="f842a-418">&lt;optional&gt;</span></span>|<span data-ttu-id="f842a-419">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="f842a-419">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f842a-420">要件</span><span class="sxs-lookup"><span data-stu-id="f842a-420">Requirements</span></span>

|<span data-ttu-id="f842a-421">要件</span><span class="sxs-lookup"><span data-stu-id="f842a-421">Requirement</span></span>| <span data-ttu-id="f842a-422">値</span><span class="sxs-lookup"><span data-stu-id="f842a-422">Value</span></span>|
|---|---|
|[<span data-ttu-id="f842a-423">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f842a-423">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f842a-424">1.0</span><span class="sxs-lookup"><span data-stu-id="f842a-424">1.0</span></span>|
|[<span data-ttu-id="f842a-425">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f842a-425">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f842a-426">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="f842a-426">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="f842a-427">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f842a-427">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f842a-428">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f842a-428">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f842a-429">例</span><span class="sxs-lookup"><span data-stu-id="f842a-429">Example</span></span>

<span data-ttu-id="f842a-430">次の例は、`makeEwsRequestAsync` を呼び出し、`GetItem` 操作を使って項目の件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="f842a-430">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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
