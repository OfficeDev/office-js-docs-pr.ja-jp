---
title: Office. メールボックス要件セット1.1
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 6bbb79af22730dd10ca2c5f1e56a96fb102ab137
ms.sourcegitcommit: 3f5d7f4794e3d3c8bc3a79fa05c54157613b9376
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/02/2019
ms.locfileid: "36064425"
---
# <a name="mailbox"></a><span data-ttu-id="cc4b2-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="cc4b2-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="cc4b2-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="cc4b2-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="cc4b2-104">Microsoft Outlook の Outlook アドインオブジェクトモデルへのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="cc4b2-105">要件</span><span class="sxs-lookup"><span data-stu-id="cc4b2-105">Requirements</span></span>

|<span data-ttu-id="cc4b2-106">要件</span><span class="sxs-lookup"><span data-stu-id="cc4b2-106">Requirement</span></span>| <span data-ttu-id="cc4b2-107">値</span><span class="sxs-lookup"><span data-stu-id="cc4b2-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc4b2-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cc4b2-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cc4b2-109">1.0</span><span class="sxs-lookup"><span data-stu-id="cc4b2-109">1.0</span></span>|
|[<span data-ttu-id="cc4b2-110">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cc4b2-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cc4b2-111">制限あり</span><span class="sxs-lookup"><span data-stu-id="cc4b2-111">Restricted</span></span>|
|[<span data-ttu-id="cc4b2-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cc4b2-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cc4b2-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="cc4b2-113">Compose or Read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="cc4b2-114">名前空間</span><span class="sxs-lookup"><span data-stu-id="cc4b2-114">Namespaces</span></span>

<span data-ttu-id="cc4b2-115">[diagnostics](Office.context.mailbox.diagnostics.md):Outlook アドインに診断情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-115">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="cc4b2-116">[item](Office.context.mailbox.item.md):Outlook アドインのメッセージや予定にアクセスするためのメソッドとプロパティを提供します。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-116">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="cc4b2-117">[userProfile](Office.context.mailbox.userProfile.md):Outlook アドインのユーザーに関する情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-117">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="cc4b2-118">メンバー</span><span class="sxs-lookup"><span data-stu-id="cc4b2-118">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="cc4b2-119">ewsUrl: String</span><span class="sxs-lookup"><span data-stu-id="cc4b2-119">ewsUrl: String</span></span>

<span data-ttu-id="cc4b2-120">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account.</span><span class="sxs-lookup"><span data-stu-id="cc4b2-120">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account.</span></span> <span data-ttu-id="cc4b2-121">Read mode only.</span><span class="sxs-lookup"><span data-stu-id="cc4b2-121">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="cc4b2-122">このメンバーは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-122">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="cc4b2-p102">
  `ewsUrl\` 値は、リモート サービスで、ユーザーのメールボックスに EWS 呼び出しを行うために使うことができます。たとえば、[選択したアイテムから添付ファイルを取得する](/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

##### <a name="type"></a><span data-ttu-id="cc4b2-125">型</span><span class="sxs-lookup"><span data-stu-id="cc4b2-125">Type</span></span>

*   <span data-ttu-id="cc4b2-126">String</span><span class="sxs-lookup"><span data-stu-id="cc4b2-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="cc4b2-127">要件</span><span class="sxs-lookup"><span data-stu-id="cc4b2-127">Requirements</span></span>

|<span data-ttu-id="cc4b2-128">要件</span><span class="sxs-lookup"><span data-stu-id="cc4b2-128">Requirement</span></span>| <span data-ttu-id="cc4b2-129">値</span><span class="sxs-lookup"><span data-stu-id="cc4b2-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc4b2-130">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cc4b2-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cc4b2-131">1.0</span><span class="sxs-lookup"><span data-stu-id="cc4b2-131">1.0</span></span>|
|[<span data-ttu-id="cc4b2-132">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cc4b2-132">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cc4b2-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cc4b2-133">ReadItem</span></span>|
|[<span data-ttu-id="cc4b2-134">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cc4b2-134">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cc4b2-135">読み取り</span><span class="sxs-lookup"><span data-stu-id="cc4b2-135">Read</span></span>|

### <a name="methods"></a><span data-ttu-id="cc4b2-136">メソッド</span><span class="sxs-lookup"><span data-stu-id="cc4b2-136">Methods</span></span>

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-11"></a><span data-ttu-id="cc4b2-137">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.1)}</span><span class="sxs-lookup"><span data-stu-id="cc4b2-137">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.1)}</span></span>

<span data-ttu-id="cc4b2-138">クライアントのローカル時間で時間情報が含まれている辞書を取得します。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-138">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="cc4b2-139">デスクトップまたは web 上の Outlook 用メールアプリは、日付と時刻に異なるタイムゾーンを使用できます。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-139">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times.</span></span> <span data-ttu-id="cc4b2-140">デスクトップ上の Outlook では、クライアントコンピューターのタイムゾーンが使用されます。Outlook on the web では、Exchange 管理センター (EAC) で設定されているタイムゾーンが使用されます。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-140">Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC).</span></span> <span data-ttu-id="cc4b2-141">日付と時刻の値を処理して、ユーザーインターフェイスに表示される値が、ユーザーが期待するタイムゾーンに常に一致するようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-141">You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="cc4b2-142">デスクトップクライアント上の Outlook でメールアプリが実行されている`convertToLocalClientTime`場合、このメソッドは、クライアントコンピューターのタイムゾーンに設定された値を持つ dictionary オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-142">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone.</span></span> <span data-ttu-id="cc4b2-143">メールアプリが web 上の Outlook で実行されている`convertToLocalClientTime`場合、このメソッドは、EAC で指定されたタイムゾーンに設定された値を持つ dictionary オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-143">If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cc4b2-144">パラメーター</span><span class="sxs-lookup"><span data-stu-id="cc4b2-144">Parameters</span></span>

|<span data-ttu-id="cc4b2-145">名前</span><span class="sxs-lookup"><span data-stu-id="cc4b2-145">Name</span></span>| <span data-ttu-id="cc4b2-146">種類</span><span class="sxs-lookup"><span data-stu-id="cc4b2-146">Type</span></span>| <span data-ttu-id="cc4b2-147">説明</span><span class="sxs-lookup"><span data-stu-id="cc4b2-147">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="cc4b2-148">Date</span><span class="sxs-lookup"><span data-stu-id="cc4b2-148">Date</span></span>|<span data-ttu-id="cc4b2-149">日付オブジェクト</span><span class="sxs-lookup"><span data-stu-id="cc4b2-149">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cc4b2-150">要件</span><span class="sxs-lookup"><span data-stu-id="cc4b2-150">Requirements</span></span>

|<span data-ttu-id="cc4b2-151">要件</span><span class="sxs-lookup"><span data-stu-id="cc4b2-151">Requirement</span></span>| <span data-ttu-id="cc4b2-152">値</span><span class="sxs-lookup"><span data-stu-id="cc4b2-152">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc4b2-153">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cc4b2-153">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cc4b2-154">1.0</span><span class="sxs-lookup"><span data-stu-id="cc4b2-154">1.0</span></span>|
|[<span data-ttu-id="cc4b2-155">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cc4b2-155">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cc4b2-156">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cc4b2-156">ReadItem</span></span>|
|[<span data-ttu-id="cc4b2-157">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cc4b2-157">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cc4b2-158">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="cc4b2-158">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="cc4b2-159">戻り値:</span><span class="sxs-lookup"><span data-stu-id="cc4b2-159">Returns:</span></span>

<span data-ttu-id="cc4b2-160">型:[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="cc4b2-160">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.1)</span></span>

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="cc4b2-161">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="cc4b2-161">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="cc4b2-162">時間情報が含まれているディクショナリから日付オブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-162">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="cc4b2-163">`convertToUtcClientTime` メソッドは、ローカルの日付と時刻を含むディクショナリを、ローカルの日付と時刻の正しい値を持つ日付オブジェクトに変換します。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-163">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cc4b2-164">パラメーター</span><span class="sxs-lookup"><span data-stu-id="cc4b2-164">Parameters</span></span>

|<span data-ttu-id="cc4b2-165">名前</span><span class="sxs-lookup"><span data-stu-id="cc4b2-165">Name</span></span>| <span data-ttu-id="cc4b2-166">種類</span><span class="sxs-lookup"><span data-stu-id="cc4b2-166">Type</span></span>| <span data-ttu-id="cc4b2-167">説明</span><span class="sxs-lookup"><span data-stu-id="cc4b2-167">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="cc4b2-168">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="cc4b2-168">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.1)|<span data-ttu-id="cc4b2-169">変換するローカル時刻の値。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-169">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cc4b2-170">要件</span><span class="sxs-lookup"><span data-stu-id="cc4b2-170">Requirements</span></span>

|<span data-ttu-id="cc4b2-171">要件</span><span class="sxs-lookup"><span data-stu-id="cc4b2-171">Requirement</span></span>| <span data-ttu-id="cc4b2-172">値</span><span class="sxs-lookup"><span data-stu-id="cc4b2-172">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc4b2-173">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cc4b2-173">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cc4b2-174">1.0</span><span class="sxs-lookup"><span data-stu-id="cc4b2-174">1.0</span></span>|
|[<span data-ttu-id="cc4b2-175">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cc4b2-175">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cc4b2-176">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cc4b2-176">ReadItem</span></span>|
|[<span data-ttu-id="cc4b2-177">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cc4b2-177">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cc4b2-178">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="cc4b2-178">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="cc4b2-179">戻り値:</span><span class="sxs-lookup"><span data-stu-id="cc4b2-179">Returns:</span></span>

<span data-ttu-id="cc4b2-180">時間が UTC で表現された日付オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-180">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="cc4b2-181">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="cc4b2-181">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="cc4b2-182">日付</span><span class="sxs-lookup"><span data-stu-id="cc4b2-182">Date</span></span></dd>

</dl>

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="cc4b2-183">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="cc4b2-183">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="cc4b2-184">既存の予定を表示します。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-184">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="cc4b2-185">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-185">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="cc4b2-186">`displayAppointmentForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存の予定を開きます。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-186">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="cc4b2-187">Outlook on the Mac では、このメソッドを使用して、定期的なアイテムの一部ではない単一の予定を表示したり、定期的なアイテムのマスター予定を表示したりすることはできませんが、一連のインスタンスを表示することはできません。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-187">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series.</span></span> <span data-ttu-id="cc4b2-188">これは、Mac 上の Outlook では、定期的なアイテムのインスタンスのプロパティ (アイテム ID を含む) にアクセスできないためです。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-188">This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="cc4b2-189">Web 上の Outlook では、フォームの本文が 32 KB 以下の文字である場合にのみ、このメソッドは指定されたフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-189">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="cc4b2-190">指定のアイテム識別子が既存の予定を表していない場合は、クライアント コンピューターまたはデバイスで空のウィンドウが開き、エラー メッセージは返されません。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-190">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cc4b2-191">パラメーター</span><span class="sxs-lookup"><span data-stu-id="cc4b2-191">Parameters</span></span>

|<span data-ttu-id="cc4b2-192">名前</span><span class="sxs-lookup"><span data-stu-id="cc4b2-192">Name</span></span>| <span data-ttu-id="cc4b2-193">種類</span><span class="sxs-lookup"><span data-stu-id="cc4b2-193">Type</span></span>| <span data-ttu-id="cc4b2-194">説明</span><span class="sxs-lookup"><span data-stu-id="cc4b2-194">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="cc4b2-195">String</span><span class="sxs-lookup"><span data-stu-id="cc4b2-195">String</span></span>|<span data-ttu-id="cc4b2-196">既存の予定の Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-196">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cc4b2-197">要件</span><span class="sxs-lookup"><span data-stu-id="cc4b2-197">Requirements</span></span>

|<span data-ttu-id="cc4b2-198">要件</span><span class="sxs-lookup"><span data-stu-id="cc4b2-198">Requirement</span></span>| <span data-ttu-id="cc4b2-199">値</span><span class="sxs-lookup"><span data-stu-id="cc4b2-199">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc4b2-200">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cc4b2-200">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cc4b2-201">1.0</span><span class="sxs-lookup"><span data-stu-id="cc4b2-201">1.0</span></span>|
|[<span data-ttu-id="cc4b2-202">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cc4b2-202">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cc4b2-203">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cc4b2-203">ReadItem</span></span>|
|[<span data-ttu-id="cc4b2-204">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cc4b2-204">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cc4b2-205">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="cc4b2-205">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cc4b2-206">例</span><span class="sxs-lookup"><span data-stu-id="cc4b2-206">Example</span></span>

```javascript
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

#### <a name="displaymessageformitemid"></a><span data-ttu-id="cc4b2-207">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="cc4b2-207">displayMessageForm(itemId)</span></span>

<span data-ttu-id="cc4b2-208">既存のメッセージを表示します。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-208">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="cc4b2-209">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-209">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="cc4b2-210">`displayMessageForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存のメッセージを開きます。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-210">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="cc4b2-211">Web 上の Outlook では、フォームの本文が 32 KB の文字数以下の場合にのみ、このメソッドは指定されたフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-211">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="cc4b2-212">指定のアイテム識別子が既存のメッセージを表していない場合は、クライアント コンピューターにはメッセージは表示されず、エラー メッセージも返されません。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-212">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="cc4b2-p106">予定を表す `itemId` を含む `displayMessageForm` を使用しないでください。既存の予定を表示するには、`displayAppointmentForm` メソッドを使用します。新しい予定を作成するフォームを表示するには、`displayNewAppointmentForm` メソッドを使用します。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-p106">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cc4b2-215">パラメーター</span><span class="sxs-lookup"><span data-stu-id="cc4b2-215">Parameters</span></span>

|<span data-ttu-id="cc4b2-216">名前</span><span class="sxs-lookup"><span data-stu-id="cc4b2-216">Name</span></span>| <span data-ttu-id="cc4b2-217">型</span><span class="sxs-lookup"><span data-stu-id="cc4b2-217">Type</span></span>| <span data-ttu-id="cc4b2-218">説明</span><span class="sxs-lookup"><span data-stu-id="cc4b2-218">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="cc4b2-219">String</span><span class="sxs-lookup"><span data-stu-id="cc4b2-219">String</span></span>|<span data-ttu-id="cc4b2-220">既存のメッセージの Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-220">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cc4b2-221">要件</span><span class="sxs-lookup"><span data-stu-id="cc4b2-221">Requirements</span></span>

|<span data-ttu-id="cc4b2-222">要件</span><span class="sxs-lookup"><span data-stu-id="cc4b2-222">Requirement</span></span>| <span data-ttu-id="cc4b2-223">値</span><span class="sxs-lookup"><span data-stu-id="cc4b2-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc4b2-224">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cc4b2-224">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cc4b2-225">1.0</span><span class="sxs-lookup"><span data-stu-id="cc4b2-225">1.0</span></span>|
|[<span data-ttu-id="cc4b2-226">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cc4b2-226">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cc4b2-227">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cc4b2-227">ReadItem</span></span>|
|[<span data-ttu-id="cc4b2-228">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cc4b2-228">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cc4b2-229">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="cc4b2-229">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cc4b2-230">例</span><span class="sxs-lookup"><span data-stu-id="cc4b2-230">Example</span></span>

```javascript
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="cc4b2-231">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="cc4b2-231">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="cc4b2-232">新しい予定を作成するためのフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-232">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="cc4b2-233">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-233">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="cc4b2-p107">`displayNewAppointmentForm` メソッドを使用すると、ユーザーが新しい予定または会議を作成できるフォームが開きます。パラメーターを指定すると、予定のフォーム フィールドにパラメーターの内容が自動的に設定されます。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-p107">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="cc4b2-236">Outlook on the web およびモバイルデバイスでは、このメソッドは常に出席者フィールドを含むフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-236">In Outlook on the web and mobile devices, this method always displays a form with an attendees field.</span></span> <span data-ttu-id="cc4b2-237">入力引数として出席者を指定しないと、このメソッドにより **[保存]** ボタンのあるフォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-237">If you do not specify any attendees as input arguments, the method displays a form with a **Save** button.</span></span> <span data-ttu-id="cc4b2-238">出席者を指定した場合には、フォームにその出席者と **[送信]** ボタンが表示されます。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-238">If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="cc4b2-p109">Outlook リッチ クライアントと Outlook RT で、`requiredAttendees`、`optionalAttendees`、または `resources` パラメーターに出席者またはリソースを指定し、このメソッドを実行すると、**[送信]** ボタンがある会議フォームが表示されます。受信者を指定せずにこのメソッドを実行すると、**[保存して閉じる]** ボタンがある予定フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-p109">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="cc4b2-241">パラメーターのいずれかが指定のサイズ制限を超える場合、または不明なパラメーター名が指定されている場合は、例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-241">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cc4b2-242">パラメーター</span><span class="sxs-lookup"><span data-stu-id="cc4b2-242">Parameters</span></span>

|<span data-ttu-id="cc4b2-243">名前</span><span class="sxs-lookup"><span data-stu-id="cc4b2-243">Name</span></span>| <span data-ttu-id="cc4b2-244">型</span><span class="sxs-lookup"><span data-stu-id="cc4b2-244">Type</span></span>| <span data-ttu-id="cc4b2-245">説明</span><span class="sxs-lookup"><span data-stu-id="cc4b2-245">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="cc4b2-246">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="cc4b2-246">Object</span></span> | <span data-ttu-id="cc4b2-247">新しい予定を記述するパラメーターのディクショナリ。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-247">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="cc4b2-248">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)&gt;</span><span class="sxs-lookup"><span data-stu-id="cc4b2-248">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)&gt;</span></span> | <span data-ttu-id="cc4b2-p110">予定に必要な各出席者について、メール アドレスを含む文字列の配列、または `EmailAddressDetails` オブジェクトを含む配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-p110">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="cc4b2-251">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)&gt;</span><span class="sxs-lookup"><span data-stu-id="cc4b2-251">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)&gt;</span></span> | <span data-ttu-id="cc4b2-p111">予定の各任意出席者について、メール アドレスを含む文字列の配列、または `EmailAddressDetails` オブジェクトを含む配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-p111">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="cc4b2-254">日付</span><span class="sxs-lookup"><span data-stu-id="cc4b2-254">Date</span></span> | <span data-ttu-id="cc4b2-255">予定の開始日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-255">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="cc4b2-256">日付</span><span class="sxs-lookup"><span data-stu-id="cc4b2-256">Date</span></span> | <span data-ttu-id="cc4b2-257">予定の終了日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-257">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="cc4b2-258">String</span><span class="sxs-lookup"><span data-stu-id="cc4b2-258">String</span></span> | <span data-ttu-id="cc4b2-p112">予定の場所を含む文字列。文字列は最大 255 文字に制限されます。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-p112">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="cc4b2-261">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="cc4b2-261">Array.&lt;String&gt;</span></span> | <span data-ttu-id="cc4b2-p113">予定に必要なリソースを含む文字列の配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-p113">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="cc4b2-264">String</span><span class="sxs-lookup"><span data-stu-id="cc4b2-264">String</span></span> | <span data-ttu-id="cc4b2-p114">予定の件名を含む文字列です。文字列は最大 255 文字に制限されます。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-p114">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="cc4b2-267">String</span><span class="sxs-lookup"><span data-stu-id="cc4b2-267">String</span></span> | <span data-ttu-id="cc4b2-p115">予定の本文。本文の内容は、最大サイズが 32 KB に制限されます。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-p115">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="cc4b2-270">要件</span><span class="sxs-lookup"><span data-stu-id="cc4b2-270">Requirements</span></span>

|<span data-ttu-id="cc4b2-271">要件</span><span class="sxs-lookup"><span data-stu-id="cc4b2-271">Requirement</span></span>| <span data-ttu-id="cc4b2-272">値</span><span class="sxs-lookup"><span data-stu-id="cc4b2-272">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc4b2-273">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cc4b2-273">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cc4b2-274">1.0</span><span class="sxs-lookup"><span data-stu-id="cc4b2-274">1.0</span></span>|
|[<span data-ttu-id="cc4b2-275">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cc4b2-275">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cc4b2-276">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cc4b2-276">ReadItem</span></span>|
|[<span data-ttu-id="cc4b2-277">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cc4b2-277">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cc4b2-278">読み取り</span><span class="sxs-lookup"><span data-stu-id="cc4b2-278">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cc4b2-279">例</span><span class="sxs-lookup"><span data-stu-id="cc4b2-279">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="cc4b2-280">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="cc4b2-280">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="cc4b2-281">Exchange Server から添付ファイルやアイテムを取得するために使うトークンを含む文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-281">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="cc4b2-p116">`getCallbackTokenAsync` メソッドは、ユーザーのメールボックスをホストする Exchange Server から不透明なトークンを取得する非同期の呼び出しを行います。コールバック トークンの有効期間は 5 分です。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-p116">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="cc4b2-p117">トークンと予定の識別子またはアイテムの識別子をサードパーティ システムに渡すことができます。サードパーティ システムは、トークンをベアラー承認トークンとして使用し、Exchange Web サービス (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) または [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) 操作を呼び出して、添付ファイルまたはアイテムを返します。たとえば、[選択したアイテムから添付ファイルを取得する](/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-p117">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="cc4b2-287">アプリが \*\*\*\* メソッドを呼び出すには、アプリのマニフェスト内に `getCallbackTokenAsync` アクセス許可が指定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-287">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cc4b2-288">パラメーター</span><span class="sxs-lookup"><span data-stu-id="cc4b2-288">Parameters</span></span>

|<span data-ttu-id="cc4b2-289">名前</span><span class="sxs-lookup"><span data-stu-id="cc4b2-289">Name</span></span>| <span data-ttu-id="cc4b2-290">型</span><span class="sxs-lookup"><span data-stu-id="cc4b2-290">Type</span></span>| <span data-ttu-id="cc4b2-291">属性</span><span class="sxs-lookup"><span data-stu-id="cc4b2-291">Attributes</span></span>| <span data-ttu-id="cc4b2-292">説明</span><span class="sxs-lookup"><span data-stu-id="cc4b2-292">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="cc4b2-293">function</span><span class="sxs-lookup"><span data-stu-id="cc4b2-293">function</span></span>||<span data-ttu-id="cc4b2-294">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-294">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="cc4b2-295">トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-295">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="cc4b2-296">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="cc4b2-296">Object</span></span>| <span data-ttu-id="cc4b2-297">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="cc4b2-297">&lt;optional&gt;</span></span>|<span data-ttu-id="cc4b2-298">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-298">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cc4b2-299">要件</span><span class="sxs-lookup"><span data-stu-id="cc4b2-299">Requirements</span></span>

|<span data-ttu-id="cc4b2-300">要件</span><span class="sxs-lookup"><span data-stu-id="cc4b2-300">Requirement</span></span>| <span data-ttu-id="cc4b2-301">値</span><span class="sxs-lookup"><span data-stu-id="cc4b2-301">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc4b2-302">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cc4b2-302">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cc4b2-303">1.0</span><span class="sxs-lookup"><span data-stu-id="cc4b2-303">1.0</span></span>|
|[<span data-ttu-id="cc4b2-304">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cc4b2-304">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cc4b2-305">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cc4b2-305">ReadItem</span></span>|
|[<span data-ttu-id="cc4b2-306">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cc4b2-306">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cc4b2-307">読み取り</span><span class="sxs-lookup"><span data-stu-id="cc4b2-307">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cc4b2-308">例</span><span class="sxs-lookup"><span data-stu-id="cc4b2-308">Example</span></span>

```javascript
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="cc4b2-309">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="cc4b2-309">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="cc4b2-310">ユーザーと Office アドインを識別するトークンを取得します。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-310">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="cc4b2-311">`getUserIdentityTokenAsync` メソッドは、[アドインとユーザーをサード パーティのシステムで識別して認証](/outlook/add-ins/authentication)することのできるトークンを返します。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-311">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="cc4b2-312">パラメーター</span><span class="sxs-lookup"><span data-stu-id="cc4b2-312">Parameters</span></span>

|<span data-ttu-id="cc4b2-313">名前</span><span class="sxs-lookup"><span data-stu-id="cc4b2-313">Name</span></span>| <span data-ttu-id="cc4b2-314">型</span><span class="sxs-lookup"><span data-stu-id="cc4b2-314">Type</span></span>| <span data-ttu-id="cc4b2-315">属性</span><span class="sxs-lookup"><span data-stu-id="cc4b2-315">Attributes</span></span>| <span data-ttu-id="cc4b2-316">説明</span><span class="sxs-lookup"><span data-stu-id="cc4b2-316">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="cc4b2-317">function</span><span class="sxs-lookup"><span data-stu-id="cc4b2-317">function</span></span>||<span data-ttu-id="cc4b2-318">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-318">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="cc4b2-319">トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-319">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="cc4b2-320">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="cc4b2-320">Object</span></span>| <span data-ttu-id="cc4b2-321">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="cc4b2-321">&lt;optional&gt;</span></span>|<span data-ttu-id="cc4b2-322">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-322">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cc4b2-323">要件</span><span class="sxs-lookup"><span data-stu-id="cc4b2-323">Requirements</span></span>

|<span data-ttu-id="cc4b2-324">要件</span><span class="sxs-lookup"><span data-stu-id="cc4b2-324">Requirement</span></span>| <span data-ttu-id="cc4b2-325">値</span><span class="sxs-lookup"><span data-stu-id="cc4b2-325">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc4b2-326">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cc4b2-326">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cc4b2-327">1.0</span><span class="sxs-lookup"><span data-stu-id="cc4b2-327">1.0</span></span>|
|[<span data-ttu-id="cc4b2-328">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cc4b2-328">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cc4b2-329">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cc4b2-329">ReadItem</span></span>|
|[<span data-ttu-id="cc4b2-330">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cc4b2-330">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cc4b2-331">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="cc4b2-331">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cc4b2-332">例</span><span class="sxs-lookup"><span data-stu-id="cc4b2-332">Example</span></span>

```javascript
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="cc4b2-333">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="cc4b2-333">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="cc4b2-334">ユーザーのメールボックスをホストしている Exchange サーバー上の Exchange Web サービス (EWS) のサービスに対して非同期の要求を行います。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-334">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="cc4b2-335">このメソッドは、次のシナリオではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-335">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="cc4b2-336">Outlook on iOS または Android</span><span class="sxs-lookup"><span data-stu-id="cc4b2-336">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="cc4b2-337">アドインが Gmail のメールボックスに読み込まれる場合</span><span class="sxs-lookup"><span data-stu-id="cc4b2-337">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="cc4b2-338">このような場合は、アドインでは [REST API を使用](/outlook/add-ins/use-rest-api)して、代わりにユーザーのメールボックスにアクセスする必要があります。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-338">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="cc4b2-p118">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange. See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span><span class="sxs-lookup"><span data-stu-id="cc4b2-p118">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange. See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="cc4b2-341">`makeEwsRequestAsync` メソッドでは、フォルダー関連アイテムを要求できません。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-341">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="cc4b2-342">XML 要求では UTF-8 エンコードを指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-342">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="cc4b2-p119">`makeEwsRequestAsync` メソッドを使用するには、アドインに **ReadWriteMailbox** アクセス許可が必要です。**ReadWriteMailbox** アクセス許可と、`makeEwsRequestAsync` メソッドで呼び出せる EWS 操作の使い方については、「[ユーザーのメールボックスへのメール アドイン アクセスのアクセス許可を指定する](/outlook/add-ins/understanding-outlook-add-in-permissions)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-p119">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="cc4b2-345">サーバー管理者は、クライアント アクセス サーバーの EWS ディレクトリで `OAuthAuthentication` を true に設定して、`makeEwsRequestAsync` メソッドで EWS 要求を行うことができるようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-345">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="cc4b2-346">バージョンの相違点</span><span class="sxs-lookup"><span data-stu-id="cc4b2-346">Version differences</span></span>

<span data-ttu-id="cc4b2-347">バージョン 15.0.4535.1004 より前のバージョンの Outlook で実行しているメール アプリで `makeEwsRequestAsync` メソッドを使う場合は、エンコード値を `ISO-8859-1` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-347">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="cc4b2-p120">Outlook on the web でメール アプリを実行している場合は、エンコード値を設定する必要はありません。mailbox.diagnostics.hostName プロパティを使って、メール アプリを Outlook で実行しているのか、Outlook on the web で実行しているのかを確認できます。mailbox.diagnostics.hostVersion プロパティを使って、どのバージョンの Outlook を使って実行しているかを確認できます。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-p120">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cc4b2-351">パラメーター</span><span class="sxs-lookup"><span data-stu-id="cc4b2-351">Parameters</span></span>

|<span data-ttu-id="cc4b2-352">名前</span><span class="sxs-lookup"><span data-stu-id="cc4b2-352">Name</span></span>| <span data-ttu-id="cc4b2-353">型</span><span class="sxs-lookup"><span data-stu-id="cc4b2-353">Type</span></span>| <span data-ttu-id="cc4b2-354">属性</span><span class="sxs-lookup"><span data-stu-id="cc4b2-354">Attributes</span></span>| <span data-ttu-id="cc4b2-355">説明</span><span class="sxs-lookup"><span data-stu-id="cc4b2-355">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="cc4b2-356">String</span><span class="sxs-lookup"><span data-stu-id="cc4b2-356">String</span></span>||<span data-ttu-id="cc4b2-357">EWS 要求です。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-357">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="cc4b2-358">function</span><span class="sxs-lookup"><span data-stu-id="cc4b2-358">function</span></span>||<span data-ttu-id="cc4b2-359">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-359">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="cc4b2-p121">The XML result of the EWS call is provided as a string in the `asyncResult.value` property. If the result exceeds 1 MB in size, an error message is returned instead.</span><span class="sxs-lookup"><span data-stu-id="cc4b2-p121">The XML result of the EWS call is provided as a string in the `asyncResult.value` property. If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="cc4b2-362">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="cc4b2-362">Object</span></span>| <span data-ttu-id="cc4b2-363">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="cc4b2-363">&lt;optional&gt;</span></span>|<span data-ttu-id="cc4b2-364">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-364">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cc4b2-365">要件</span><span class="sxs-lookup"><span data-stu-id="cc4b2-365">Requirements</span></span>

|<span data-ttu-id="cc4b2-366">要件</span><span class="sxs-lookup"><span data-stu-id="cc4b2-366">Requirement</span></span>| <span data-ttu-id="cc4b2-367">値</span><span class="sxs-lookup"><span data-stu-id="cc4b2-367">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc4b2-368">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cc4b2-368">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cc4b2-369">1.0</span><span class="sxs-lookup"><span data-stu-id="cc4b2-369">1.0</span></span>|
|[<span data-ttu-id="cc4b2-370">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cc4b2-370">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cc4b2-371">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="cc4b2-371">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="cc4b2-372">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cc4b2-372">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cc4b2-373">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="cc4b2-373">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cc4b2-374">例</span><span class="sxs-lookup"><span data-stu-id="cc4b2-374">Example</span></span>

<span data-ttu-id="cc4b2-375">次の例は、`makeEwsRequestAsync` を呼び出し、`GetItem` 操作を使って項目の件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="cc4b2-375">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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
