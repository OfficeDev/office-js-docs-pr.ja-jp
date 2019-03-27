---
title: Office. メールボックス要件セット1.1
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 629d0e5cde637ef209736dd9359ea59d0f6e0e47
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870605"
---
# <a name="mailbox"></a><span data-ttu-id="dcc28-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="dcc28-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="dcc28-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="dcc28-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="dcc28-104">Microsoft Outlook と Microsoft Outlook on the web の Outlook アドイン オブジェクト モデルへのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="dcc28-104">Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="dcc28-105">要件</span><span class="sxs-lookup"><span data-stu-id="dcc28-105">Requirements</span></span>

|<span data-ttu-id="dcc28-106">要件</span><span class="sxs-lookup"><span data-stu-id="dcc28-106">Requirement</span></span>| <span data-ttu-id="dcc28-107">値</span><span class="sxs-lookup"><span data-stu-id="dcc28-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="dcc28-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dcc28-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dcc28-109">1.0</span><span class="sxs-lookup"><span data-stu-id="dcc28-109">1.0</span></span>|
|[<span data-ttu-id="dcc28-110">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="dcc28-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dcc28-111">制限あり</span><span class="sxs-lookup"><span data-stu-id="dcc28-111">Restricted</span></span>|
|[<span data-ttu-id="dcc28-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dcc28-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dcc28-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="dcc28-113">Compose or Read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="dcc28-114">名前空間</span><span class="sxs-lookup"><span data-stu-id="dcc28-114">Namespaces</span></span>

<span data-ttu-id="dcc28-115">[diagnostics](Office.context.mailbox.diagnostics.md): Outlook アドインに診断情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="dcc28-115">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="dcc28-116">[item](Office.context.mailbox.item.md):Outlook アドインのメッセージや予定にアクセスするためのメソッドとプロパティを提供します。</span><span class="sxs-lookup"><span data-stu-id="dcc28-116">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="dcc28-117">[userProfile](Office.context.mailbox.userProfile.md):Outlook アドインのユーザーに関する情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="dcc28-117">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="dcc28-118">メンバー</span><span class="sxs-lookup"><span data-stu-id="dcc28-118">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="dcc28-119">ewsUrl: 文字列</span><span class="sxs-lookup"><span data-stu-id="dcc28-119">ewsUrl :String</span></span>

<span data-ttu-id="dcc28-p101">このメール アカウントの Exchange Web サービス (EWS) エンドポイントの URL を取得します。読み取りモードのみです。</span><span class="sxs-lookup"><span data-stu-id="dcc28-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="dcc28-122">このメンバーは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="dcc28-122">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="dcc28-p102">`ewsUrl` 値は、リモート サービスで、ユーザーのメールボックスに EWS 呼び出しを行うために使用することができます。たとえば、[選択した項目から添付ファイルを取得する](/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="dcc28-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

##### <a name="type"></a><span data-ttu-id="dcc28-125">型</span><span class="sxs-lookup"><span data-stu-id="dcc28-125">Type</span></span>

*   <span data-ttu-id="dcc28-126">String</span><span class="sxs-lookup"><span data-stu-id="dcc28-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="dcc28-127">要件</span><span class="sxs-lookup"><span data-stu-id="dcc28-127">Requirements</span></span>

|<span data-ttu-id="dcc28-128">要件</span><span class="sxs-lookup"><span data-stu-id="dcc28-128">Requirement</span></span>| <span data-ttu-id="dcc28-129">値</span><span class="sxs-lookup"><span data-stu-id="dcc28-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="dcc28-130">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dcc28-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dcc28-131">1.0</span><span class="sxs-lookup"><span data-stu-id="dcc28-131">1.0</span></span>|
|[<span data-ttu-id="dcc28-132">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="dcc28-132">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dcc28-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dcc28-133">ReadItem</span></span>|
|[<span data-ttu-id="dcc28-134">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dcc28-134">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dcc28-135">読み取り</span><span class="sxs-lookup"><span data-stu-id="dcc28-135">Read</span></span>|

### <a name="methods"></a><span data-ttu-id="dcc28-136">メソッド</span><span class="sxs-lookup"><span data-stu-id="dcc28-136">Methods</span></span>

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime"></a><span data-ttu-id="dcc28-137">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_1/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="dcc28-137">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_1/office.LocalClientTime)}</span></span>

<span data-ttu-id="dcc28-138">クライアントの現地時間の時間情報が含まれている辞書を取得します。</span><span class="sxs-lookup"><span data-stu-id="dcc28-138">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="dcc28-p103">Outlook 用メール アプリや Outlook Web App で使う日付と時刻は、異なるタイム ゾーンを使うことができます。Outlook では、クライアント コンピューターのタイム ゾーンを使います。Outlook Web App では、Exchange 管理センター (EAC) で設定されたタイム ゾーンを使います。ユーザー インターフェイスに表示される値が、常にユーザーが期待するタイム ゾーンと一致するように日付と時刻の値を処理する必要があります。</span><span class="sxs-lookup"><span data-stu-id="dcc28-p103">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="dcc28-p104">Outlook でメール アプリが実行されている場合、`convertToLocalClientTime` メソッドは、クライアント コンピューターのタイム ゾーンに設定された値で辞書オブジェクトを返します。Outlook Web Apps でメール アプリが実行されている場合、`convertToLocalClientTime` メソッドは、EAC で指定されたタイム ゾーンに設定された値で辞書オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="dcc28-p104">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="dcc28-144">パラメーター</span><span class="sxs-lookup"><span data-stu-id="dcc28-144">Parameters</span></span>

|<span data-ttu-id="dcc28-145">名前</span><span class="sxs-lookup"><span data-stu-id="dcc28-145">Name</span></span>| <span data-ttu-id="dcc28-146">種類</span><span class="sxs-lookup"><span data-stu-id="dcc28-146">Type</span></span>| <span data-ttu-id="dcc28-147">説明</span><span class="sxs-lookup"><span data-stu-id="dcc28-147">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="dcc28-148">Date</span><span class="sxs-lookup"><span data-stu-id="dcc28-148">Date</span></span>|<span data-ttu-id="dcc28-149">Date オブジェクト</span><span class="sxs-lookup"><span data-stu-id="dcc28-149">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dcc28-150">要件</span><span class="sxs-lookup"><span data-stu-id="dcc28-150">Requirements</span></span>

|<span data-ttu-id="dcc28-151">要件</span><span class="sxs-lookup"><span data-stu-id="dcc28-151">Requirement</span></span>| <span data-ttu-id="dcc28-152">値</span><span class="sxs-lookup"><span data-stu-id="dcc28-152">Value</span></span>|
|---|---|
|[<span data-ttu-id="dcc28-153">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dcc28-153">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dcc28-154">1.0</span><span class="sxs-lookup"><span data-stu-id="dcc28-154">1.0</span></span>|
|[<span data-ttu-id="dcc28-155">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="dcc28-155">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dcc28-156">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dcc28-156">ReadItem</span></span>|
|[<span data-ttu-id="dcc28-157">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dcc28-157">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dcc28-158">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="dcc28-158">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="dcc28-159">戻り値:</span><span class="sxs-lookup"><span data-stu-id="dcc28-159">Returns:</span></span>

<span data-ttu-id="dcc28-160">種類:[LocalClientTime](/javascript/api/outlook_1_1/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="dcc28-160">Type: [LocalClientTime](/javascript/api/outlook_1_1/office.LocalClientTime)</span></span>

####  <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="dcc28-161">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="dcc28-161">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="dcc28-162">時間情報が含まれている辞書から Date オブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="dcc28-162">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="dcc28-163">`convertToUtcClientTime` メソッドは、ローカルの日付と時刻が含まれる辞書を、ローカルの日付と時刻の正しい値をもつ Date オブジェクトに変換します。</span><span class="sxs-lookup"><span data-stu-id="dcc28-163">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="dcc28-164">パラメーター</span><span class="sxs-lookup"><span data-stu-id="dcc28-164">Parameters</span></span>

|<span data-ttu-id="dcc28-165">名前</span><span class="sxs-lookup"><span data-stu-id="dcc28-165">Name</span></span>| <span data-ttu-id="dcc28-166">種類</span><span class="sxs-lookup"><span data-stu-id="dcc28-166">Type</span></span>| <span data-ttu-id="dcc28-167">説明</span><span class="sxs-lookup"><span data-stu-id="dcc28-167">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="dcc28-168">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="dcc28-168">LocalClientTime</span></span>](/javascript/api/outlook_1_1/office.LocalClientTime)|<span data-ttu-id="dcc28-169">変換するローカル時刻の値。</span><span class="sxs-lookup"><span data-stu-id="dcc28-169">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dcc28-170">要件</span><span class="sxs-lookup"><span data-stu-id="dcc28-170">Requirements</span></span>

|<span data-ttu-id="dcc28-171">要件</span><span class="sxs-lookup"><span data-stu-id="dcc28-171">Requirement</span></span>| <span data-ttu-id="dcc28-172">値</span><span class="sxs-lookup"><span data-stu-id="dcc28-172">Value</span></span>|
|---|---|
|[<span data-ttu-id="dcc28-173">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dcc28-173">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dcc28-174">1.0</span><span class="sxs-lookup"><span data-stu-id="dcc28-174">1.0</span></span>|
|[<span data-ttu-id="dcc28-175">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="dcc28-175">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dcc28-176">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dcc28-176">ReadItem</span></span>|
|[<span data-ttu-id="dcc28-177">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dcc28-177">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dcc28-178">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="dcc28-178">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="dcc28-179">戻り値 :</span><span class="sxs-lookup"><span data-stu-id="dcc28-179">Returns:</span></span>

<span data-ttu-id="dcc28-180">時間が UTC で表現された Date オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="dcc28-180">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="dcc28-181">

<dt>種類</dt>

</span><span class="sxs-lookup"><span data-stu-id="dcc28-181">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="dcc28-182">日付</span><span class="sxs-lookup"><span data-stu-id="dcc28-182">Date</span></span></dd>

</dl>

####  <a name="displayappointmentformitemid"></a><span data-ttu-id="dcc28-183">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="dcc28-183">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="dcc28-184">既存の予定を表示します。</span><span class="sxs-lookup"><span data-stu-id="dcc28-184">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="dcc28-185">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="dcc28-185">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="dcc28-186">`displayAppointmentForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存の予定を開きます。</span><span class="sxs-lookup"><span data-stu-id="dcc28-186">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="dcc28-p105">Outlook for Mac では、このメソッドを使用して、定期的に繰り返される予定の一部ではない単発の予定、または定期的に繰り替えされる予定の元の予定を表示できます。ただし、一連の予定のインスタンスは表示できません。これは、Outlook for Mac では、定期的に繰り返されるインスタンスのプロパティ  (項目 ID を含む) にアクセスできないためです。</span><span class="sxs-lookup"><span data-stu-id="dcc28-p105">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="dcc28-189">Outlook Web App では、このメソッドで、指定されたフォームの本文が 32 KB 以下の文字数の場合に、そのフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="dcc28-189">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="dcc28-190">指定した項目識別子が既存の予定を識別しない場合には、クライアント コンピュータまたはデバイスで空白のウィンドウが開き、エラー メッセージは返されません。</span><span class="sxs-lookup"><span data-stu-id="dcc28-190">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="dcc28-191">パラメーター</span><span class="sxs-lookup"><span data-stu-id="dcc28-191">Parameters</span></span>

|<span data-ttu-id="dcc28-192">名前</span><span class="sxs-lookup"><span data-stu-id="dcc28-192">Name</span></span>| <span data-ttu-id="dcc28-193">種類</span><span class="sxs-lookup"><span data-stu-id="dcc28-193">Type</span></span>| <span data-ttu-id="dcc28-194">説明</span><span class="sxs-lookup"><span data-stu-id="dcc28-194">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="dcc28-195">文字列</span><span class="sxs-lookup"><span data-stu-id="dcc28-195">String</span></span>|<span data-ttu-id="dcc28-196">既存の予定表の予定の Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="dcc28-196">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dcc28-197">要件</span><span class="sxs-lookup"><span data-stu-id="dcc28-197">Requirements</span></span>

|<span data-ttu-id="dcc28-198">要件</span><span class="sxs-lookup"><span data-stu-id="dcc28-198">Requirement</span></span>| <span data-ttu-id="dcc28-199">値</span><span class="sxs-lookup"><span data-stu-id="dcc28-199">Value</span></span>|
|---|---|
|[<span data-ttu-id="dcc28-200">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dcc28-200">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dcc28-201">1.0</span><span class="sxs-lookup"><span data-stu-id="dcc28-201">1.0</span></span>|
|[<span data-ttu-id="dcc28-202">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="dcc28-202">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dcc28-203">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dcc28-203">ReadItem</span></span>|
|[<span data-ttu-id="dcc28-204">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dcc28-204">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dcc28-205">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="dcc28-205">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dcc28-206">例</span><span class="sxs-lookup"><span data-stu-id="dcc28-206">Example</span></span>

```javascript
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  <a name="displaymessageformitemid"></a><span data-ttu-id="dcc28-207">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="dcc28-207">displayMessageForm(itemId)</span></span>

<span data-ttu-id="dcc28-208">既存のメッセージを表示します。</span><span class="sxs-lookup"><span data-stu-id="dcc28-208">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="dcc28-209">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="dcc28-209">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="dcc28-210">`displayMessageForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存のメッセージを開きます。</span><span class="sxs-lookup"><span data-stu-id="dcc28-210">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="dcc28-211">Outlook Web App では、このメソッドは、指定されたフォームの本文が 32 KB 以下の文字数の場合に、そのフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="dcc28-211">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="dcc28-212">指定した項目識別子が既存のメッセージを識別しない場合には、クラアント コンピュータでメッセージは表示されず、エラー メッセージは返されません。</span><span class="sxs-lookup"><span data-stu-id="dcc28-212">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="dcc28-p106">予定を表す `itemId` が含まれる `displayMessageForm` を使用しないでください。`displayAppointmentForm` メソッドを使用して既存の予定を表示し、`displayNewAppointmentForm` を使用して新しい予定を作成するフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="dcc28-p106">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="dcc28-215">パラメーター</span><span class="sxs-lookup"><span data-stu-id="dcc28-215">Parameters</span></span>

|<span data-ttu-id="dcc28-216">名前</span><span class="sxs-lookup"><span data-stu-id="dcc28-216">Name</span></span>| <span data-ttu-id="dcc28-217">型</span><span class="sxs-lookup"><span data-stu-id="dcc28-217">Type</span></span>| <span data-ttu-id="dcc28-218">説明</span><span class="sxs-lookup"><span data-stu-id="dcc28-218">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="dcc28-219">String</span><span class="sxs-lookup"><span data-stu-id="dcc28-219">String</span></span>|<span data-ttu-id="dcc28-220">既存のメッセージの Exchange Web サービス(EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="dcc28-220">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dcc28-221">要件</span><span class="sxs-lookup"><span data-stu-id="dcc28-221">Requirements</span></span>

|<span data-ttu-id="dcc28-222">要件</span><span class="sxs-lookup"><span data-stu-id="dcc28-222">Requirement</span></span>| <span data-ttu-id="dcc28-223">値</span><span class="sxs-lookup"><span data-stu-id="dcc28-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="dcc28-224">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dcc28-224">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dcc28-225">1.0</span><span class="sxs-lookup"><span data-stu-id="dcc28-225">1.0</span></span>|
|[<span data-ttu-id="dcc28-226">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="dcc28-226">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dcc28-227">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dcc28-227">ReadItem</span></span>|
|[<span data-ttu-id="dcc28-228">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dcc28-228">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dcc28-229">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="dcc28-229">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dcc28-230">例</span><span class="sxs-lookup"><span data-stu-id="dcc28-230">Example</span></span>

```javascript
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="dcc28-231">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="dcc28-231">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="dcc28-232">新しい予定を作成するためのフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="dcc28-232">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="dcc28-233">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="dcc28-233">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="dcc28-p107">`displayNewAppointmentForm` メソッドを使用すると、ユーザーが新しい予定または会議を作成できるフォームが開きます。パラメータを指定すると、予定のフォーム フィールドにパラメータの内容が自動的に入力されます。</span><span class="sxs-lookup"><span data-stu-id="dcc28-p107">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="dcc28-p108">Outlook Web App および OWA for Devices では、このメソッドは出席者フィールドが含まれるフォームを常に表示します。入力因数として出席者を指定しない場合には、このメソッドは [**保存**] ボタンのあるフォームを表示します。出席者を指定した場合には、フォームにはその出席者と [**送信**] ボタンが含まれます。</span><span class="sxs-lookup"><span data-stu-id="dcc28-p108">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="dcc28-p109">Outlook リッチ クライアントおよび Outlook RT では、`requiredAttendees`、`optionalAttendees` または `resources` パラメータに出席者またはリソースを指定した場合、このメソッドは [**送信**] ボタンがある会議フォームを表示します。受信者を指定しない場合には、このメソッドは [**保存して閉じる**] ボタンがある予定フォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="dcc28-p109">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="dcc28-241">パラメータのいずれかが指定されたサイズ制限を超えた場合、または不明なパラメータ名が指定された場合には、例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="dcc28-241">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="dcc28-242">パラメーター</span><span class="sxs-lookup"><span data-stu-id="dcc28-242">Parameters</span></span>

|<span data-ttu-id="dcc28-243">名前</span><span class="sxs-lookup"><span data-stu-id="dcc28-243">Name</span></span>| <span data-ttu-id="dcc28-244">型</span><span class="sxs-lookup"><span data-stu-id="dcc28-244">Type</span></span>| <span data-ttu-id="dcc28-245">説明</span><span class="sxs-lookup"><span data-stu-id="dcc28-245">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="dcc28-246">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="dcc28-246">Object</span></span> | <span data-ttu-id="dcc28-247">新しい予定を記述するパラメータの辞書。</span><span class="sxs-lookup"><span data-stu-id="dcc28-247">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="dcc28-248">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="dcc28-248">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="dcc28-p110">予定への各必須出席者の電子メール アドレスを含む文字列の配列、または出席者の `EmailAddressDetails` オブジェクトを含む配列。この配列は、最大 100 エントリまでに制限されています。</span><span class="sxs-lookup"><span data-stu-id="dcc28-p110">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="dcc28-251">配列。&lt;文字列&gt; | 配列。&lt;[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="dcc28-251">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="dcc28-p111">予定への各任意出席者の電子メール アドレスを含む文字列の配列、または出席者の `EmailAddressDetails` オブジェクトを含む配列。この配列は、最大 100 エントリまでに制限されています。</span><span class="sxs-lookup"><span data-stu-id="dcc28-p111">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="dcc28-254">日付</span><span class="sxs-lookup"><span data-stu-id="dcc28-254">Date</span></span> | <span data-ttu-id="dcc28-255">予定の開始日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="dcc28-255">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="dcc28-256">日付</span><span class="sxs-lookup"><span data-stu-id="dcc28-256">Date</span></span> | <span data-ttu-id="dcc28-257">予定の終了日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="dcc28-257">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="dcc28-258">String</span><span class="sxs-lookup"><span data-stu-id="dcc28-258">String</span></span> | <span data-ttu-id="dcc28-p112">予定の場所を含む文字列。文字列は最大 255 文字までに制限されています。</span><span class="sxs-lookup"><span data-stu-id="dcc28-p112">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="dcc28-261">配列。&lt; 文字列&gt;</span><span class="sxs-lookup"><span data-stu-id="dcc28-261">Array.&lt;String&gt;</span></span> | <span data-ttu-id="dcc28-p113">予定に必要なリソースを含む文字列の配列。配列は最大 100 エントリまでに制限されています。</span><span class="sxs-lookup"><span data-stu-id="dcc28-p113">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="dcc28-264">String</span><span class="sxs-lookup"><span data-stu-id="dcc28-264">String</span></span> | <span data-ttu-id="dcc28-p114">予定の件名を含む文字列。文字列は最大 255 文字までに制限されています。</span><span class="sxs-lookup"><span data-stu-id="dcc28-p114">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="dcc28-267">文字列</span><span class="sxs-lookup"><span data-stu-id="dcc28-267">String</span></span> | <span data-ttu-id="dcc28-p115">予定の本文。本文の内容は、最大サイズが 32 KB までに制限されています。</span><span class="sxs-lookup"><span data-stu-id="dcc28-p115">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="dcc28-270">要件</span><span class="sxs-lookup"><span data-stu-id="dcc28-270">Requirements</span></span>

|<span data-ttu-id="dcc28-271">要件</span><span class="sxs-lookup"><span data-stu-id="dcc28-271">Requirement</span></span>| <span data-ttu-id="dcc28-272">値</span><span class="sxs-lookup"><span data-stu-id="dcc28-272">Value</span></span>|
|---|---|
|[<span data-ttu-id="dcc28-273">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dcc28-273">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dcc28-274">1.0</span><span class="sxs-lookup"><span data-stu-id="dcc28-274">1.0</span></span>|
|[<span data-ttu-id="dcc28-275">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="dcc28-275">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dcc28-276">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dcc28-276">ReadItem</span></span>|
|[<span data-ttu-id="dcc28-277">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dcc28-277">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dcc28-278">読み取り</span><span class="sxs-lookup"><span data-stu-id="dcc28-278">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dcc28-279">例</span><span class="sxs-lookup"><span data-stu-id="dcc28-279">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="dcc28-280">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="dcc28-280">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="dcc28-281">Exchange Server から添付ファイルやアイテムを取得するために使用するトークンを含む文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="dcc28-281">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="dcc28-p116">`getCallbackTokenAsync` メソッドは、非同期の呼び出しを行なって、ユーザーのメールボックスをホストする Exchange Server から opaque トークンを取得します。コールバック トークンの有効期間は 5 分です。</span><span class="sxs-lookup"><span data-stu-id="dcc28-p116">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="dcc28-p117">このトークンと、添付ファイル識別子または項目識別子は、サードパーティーのシステムに渡すことができます。サードパーティーのシステムでは、添付ファイルまたは項目を返すための Exchange Web サービス (EWS) の [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) または [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) 操作を呼び出すのに、トークンをベアラー承認トークンとして使用します。たとえば、リモート サービスを作成して[選択した項目から添付ファイルを取得](/outlook/add-ins/get-attachments-of-an-outlook-item)することができます。</span><span class="sxs-lookup"><span data-stu-id="dcc28-p117">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="dcc28-287">アプリが \*\*\*\* メソッドを呼び出すには、アプリのマニフェスト内に `getCallbackTokenAsync` アクセス許可が指定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="dcc28-287">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="dcc28-288">パラメーター</span><span class="sxs-lookup"><span data-stu-id="dcc28-288">Parameters</span></span>

|<span data-ttu-id="dcc28-289">名前</span><span class="sxs-lookup"><span data-stu-id="dcc28-289">Name</span></span>| <span data-ttu-id="dcc28-290">型</span><span class="sxs-lookup"><span data-stu-id="dcc28-290">Type</span></span>| <span data-ttu-id="dcc28-291">属性</span><span class="sxs-lookup"><span data-stu-id="dcc28-291">Attributes</span></span>| <span data-ttu-id="dcc28-292">説明</span><span class="sxs-lookup"><span data-stu-id="dcc28-292">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="dcc28-293">function</span><span class="sxs-lookup"><span data-stu-id="dcc28-293">function</span></span>||<span data-ttu-id="dcc28-294">メソッドが完了すると、`callback` パラメータに渡された関数が、単一パラメータ `asyncResult` で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="dcc28-294">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="dcc28-295">トークンは、`asyncResult.value`プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="dcc28-295">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="dcc28-296">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="dcc28-296">Object</span></span>| <span data-ttu-id="dcc28-297">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="dcc28-297">&lt;optional&gt;</span></span>|<span data-ttu-id="dcc28-298">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="dcc28-298">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dcc28-299">要件</span><span class="sxs-lookup"><span data-stu-id="dcc28-299">Requirements</span></span>

|<span data-ttu-id="dcc28-300">要件</span><span class="sxs-lookup"><span data-stu-id="dcc28-300">Requirement</span></span>| <span data-ttu-id="dcc28-301">値</span><span class="sxs-lookup"><span data-stu-id="dcc28-301">Value</span></span>|
|---|---|
|[<span data-ttu-id="dcc28-302">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dcc28-302">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dcc28-303">1.0</span><span class="sxs-lookup"><span data-stu-id="dcc28-303">1.0</span></span>|
|[<span data-ttu-id="dcc28-304">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="dcc28-304">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dcc28-305">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dcc28-305">ReadItem</span></span>|
|[<span data-ttu-id="dcc28-306">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dcc28-306">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dcc28-307">読み取り</span><span class="sxs-lookup"><span data-stu-id="dcc28-307">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dcc28-308">例</span><span class="sxs-lookup"><span data-stu-id="dcc28-308">Example</span></span>

```javascript
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="dcc28-309">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="dcc28-309">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="dcc28-310">ユーザーと Office アドインを識別するトークンを取得します。</span><span class="sxs-lookup"><span data-stu-id="dcc28-310">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="dcc28-311">`getUserIdentityTokenAsync` メソッドは、[アドインとユーザーをサードパーティのシステムで識別して認証](/outlook/add-ins/authentication)するのに使用できるトークンを返します。</span><span class="sxs-lookup"><span data-stu-id="dcc28-311">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="dcc28-312">パラメーター</span><span class="sxs-lookup"><span data-stu-id="dcc28-312">Parameters</span></span>

|<span data-ttu-id="dcc28-313">名前</span><span class="sxs-lookup"><span data-stu-id="dcc28-313">Name</span></span>| <span data-ttu-id="dcc28-314">型</span><span class="sxs-lookup"><span data-stu-id="dcc28-314">Type</span></span>| <span data-ttu-id="dcc28-315">属性</span><span class="sxs-lookup"><span data-stu-id="dcc28-315">Attributes</span></span>| <span data-ttu-id="dcc28-316">説明</span><span class="sxs-lookup"><span data-stu-id="dcc28-316">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="dcc28-317">function</span><span class="sxs-lookup"><span data-stu-id="dcc28-317">function</span></span>||<span data-ttu-id="dcc28-318">メソッドが完了すると、`callback` パラメータに渡された関数が、単一パラメータ `asyncResult` で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="dcc28-318">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="dcc28-319">トークンは、`asyncResult.value`プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="dcc28-319">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="dcc28-320">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="dcc28-320">Object</span></span>| <span data-ttu-id="dcc28-321">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="dcc28-321">&lt;optional&gt;</span></span>|<span data-ttu-id="dcc28-322">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="dcc28-322">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dcc28-323">要件</span><span class="sxs-lookup"><span data-stu-id="dcc28-323">Requirements</span></span>

|<span data-ttu-id="dcc28-324">要件</span><span class="sxs-lookup"><span data-stu-id="dcc28-324">Requirement</span></span>| <span data-ttu-id="dcc28-325">値</span><span class="sxs-lookup"><span data-stu-id="dcc28-325">Value</span></span>|
|---|---|
|[<span data-ttu-id="dcc28-326">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dcc28-326">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dcc28-327">1.0</span><span class="sxs-lookup"><span data-stu-id="dcc28-327">1.0</span></span>|
|[<span data-ttu-id="dcc28-328">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="dcc28-328">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dcc28-329">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dcc28-329">ReadItem</span></span>|
|[<span data-ttu-id="dcc28-330">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dcc28-330">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dcc28-331">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="dcc28-331">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dcc28-332">例</span><span class="sxs-lookup"><span data-stu-id="dcc28-332">Example</span></span>

```javascript
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="dcc28-333">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="dcc28-333">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="dcc28-334">ユーザーのメールボックスをホストしている Exchange サーバー上の Exchange Web サービス (EWS) のサービスに対して非同期の要求を行います。</span><span class="sxs-lookup"><span data-stu-id="dcc28-334">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="dcc28-335">このメソッドは、次のシナリオではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="dcc28-335">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="dcc28-336">Outlook for iOS または Outlook for Android で</span><span class="sxs-lookup"><span data-stu-id="dcc28-336">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="dcc28-337">アドインが Gmail のメールボックスにロードされる場合</span><span class="sxs-lookup"><span data-stu-id="dcc28-337">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="dcc28-338">これらの場合、アドインではユーザーのメールボックスにアクセスするために、代わりに [REST API を使用する](/outlook/add-ins/use-rest-api)必要があります。</span><span class="sxs-lookup"><span data-stu-id="dcc28-338">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="dcc28-p118">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange. See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span><span class="sxs-lookup"><span data-stu-id="dcc28-p118">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange. See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="dcc28-341">`makeEwsRequestAsync` メソッドでは、フォルダー関連アイテムを要求できません。</span><span class="sxs-lookup"><span data-stu-id="dcc28-341">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="dcc28-342">XML 要求では、UTF-8 エンコードを指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="dcc28-342">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="dcc28-p119">アドインには、`makeEwsRequestAsync` メソッドを使用するために **ReadWriteMailbox** アクセス許可が必要です。**ReadWriteMailbox** アクセス許可と、`makeEwsRequestAsync` メソッドで呼び出すことのできる EWS 操作の使用の詳細については、「[ユーザーのメールボックスへのメール アドイン アクセスのアクセス許可を指定する](/outlook/add-ins/understanding-outlook-add-in-permissions)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="dcc28-p119">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="dcc28-345">サーバー管理者は、クライアント アクセス サーバーの EWS ディレクトリで `OAuthAuthentication` を true に設定して、`makeEwsRequestAsync` メソッドで EWS 要求を行うことができるようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="dcc28-345">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="dcc28-346">バージョンの相違点</span><span class="sxs-lookup"><span data-stu-id="dcc28-346">Version differences</span></span>

<span data-ttu-id="dcc28-347">バージョン 15.0.4535.1004 より前のバージョンの Outlook で実行しているメール アプリで `makeEwsRequestAsync` メソッドを使用する場合には、エンコード値を `ISO-8859-1` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="dcc28-347">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="dcc28-p120">メール アプリが Outlook on the web で実行されている場合には、エンコード値を設定する必要はありません。メールボックスを使用してメール アプリが Outlook で実行されているのか、Outlook on the web で実行されているのかを判断する必要があります。mailbox.diagnostics.hostVersion プロパティを使用すれば、どのバージョンの Outlook が実行されているのかがわかります。</span><span class="sxs-lookup"><span data-stu-id="dcc28-p120">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="dcc28-351">パラメーター</span><span class="sxs-lookup"><span data-stu-id="dcc28-351">Parameters</span></span>

|<span data-ttu-id="dcc28-352">名前</span><span class="sxs-lookup"><span data-stu-id="dcc28-352">Name</span></span>| <span data-ttu-id="dcc28-353">型</span><span class="sxs-lookup"><span data-stu-id="dcc28-353">Type</span></span>| <span data-ttu-id="dcc28-354">属性</span><span class="sxs-lookup"><span data-stu-id="dcc28-354">Attributes</span></span>| <span data-ttu-id="dcc28-355">説明</span><span class="sxs-lookup"><span data-stu-id="dcc28-355">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="dcc28-356">文字列</span><span class="sxs-lookup"><span data-stu-id="dcc28-356">String</span></span>||<span data-ttu-id="dcc28-357">EWS 要求。</span><span class="sxs-lookup"><span data-stu-id="dcc28-357">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="dcc28-358">関数</span><span class="sxs-lookup"><span data-stu-id="dcc28-358">function</span></span>||<span data-ttu-id="dcc28-359">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="dcc28-359">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="dcc28-p121">The XML result of the EWS call is provided as a string in the `asyncResult.value` property. If the result exceeds 1 MB in size, an error message is returned instead.</span><span class="sxs-lookup"><span data-stu-id="dcc28-p121">The XML result of the EWS call is provided as a string in the `asyncResult.value` property. If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="dcc28-362">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="dcc28-362">Object</span></span>| <span data-ttu-id="dcc28-363">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="dcc28-363">&lt;optional&gt;</span></span>|<span data-ttu-id="dcc28-364">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="dcc28-364">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dcc28-365">要件</span><span class="sxs-lookup"><span data-stu-id="dcc28-365">Requirements</span></span>

|<span data-ttu-id="dcc28-366">要件</span><span class="sxs-lookup"><span data-stu-id="dcc28-366">Requirement</span></span>| <span data-ttu-id="dcc28-367">値</span><span class="sxs-lookup"><span data-stu-id="dcc28-367">Value</span></span>|
|---|---|
|[<span data-ttu-id="dcc28-368">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dcc28-368">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dcc28-369">1.0</span><span class="sxs-lookup"><span data-stu-id="dcc28-369">1.0</span></span>|
|[<span data-ttu-id="dcc28-370">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="dcc28-370">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dcc28-371">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="dcc28-371">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="dcc28-372">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dcc28-372">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dcc28-373">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="dcc28-373">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dcc28-374">例</span><span class="sxs-lookup"><span data-stu-id="dcc28-374">Example</span></span>

<span data-ttu-id="dcc28-375">次の例は、`makeEwsRequestAsync` を呼び出し、`GetItem` 操作を使って項目の件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="dcc28-375">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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
