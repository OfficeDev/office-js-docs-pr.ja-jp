---
title: Office. メールボックス要件セット1.2
description: ''
ms.date: 08/30/2019
localization_priority: Normal
ms.openlocfilehash: 2002b7784d0d7295762d1f692e7a0115f1f97059
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696345"
---
# <a name="mailbox"></a><span data-ttu-id="ff5ed-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="ff5ed-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="ff5ed-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="ff5ed-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="ff5ed-104">Microsoft Outlook の Outlook アドインオブジェクトモデルへのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="ff5ed-105">要件</span><span class="sxs-lookup"><span data-stu-id="ff5ed-105">Requirements</span></span>

|<span data-ttu-id="ff5ed-106">要件</span><span class="sxs-lookup"><span data-stu-id="ff5ed-106">Requirement</span></span>| <span data-ttu-id="ff5ed-107">値</span><span class="sxs-lookup"><span data-stu-id="ff5ed-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="ff5ed-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ff5ed-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ff5ed-109">1.0</span><span class="sxs-lookup"><span data-stu-id="ff5ed-109">1.0</span></span>|
|[<span data-ttu-id="ff5ed-110">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ff5ed-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ff5ed-111">制限あり</span><span class="sxs-lookup"><span data-stu-id="ff5ed-111">Restricted</span></span>|
|[<span data-ttu-id="ff5ed-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ff5ed-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ff5ed-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ff5ed-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="ff5ed-114">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="ff5ed-114">Members and methods</span></span>

| <span data-ttu-id="ff5ed-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="ff5ed-115">Member</span></span> | <span data-ttu-id="ff5ed-116">種類</span><span class="sxs-lookup"><span data-stu-id="ff5ed-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="ff5ed-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="ff5ed-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="ff5ed-118">メンバー</span><span class="sxs-lookup"><span data-stu-id="ff5ed-118">Member</span></span> |
| [<span data-ttu-id="ff5ed-119">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="ff5ed-119">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="ff5ed-120">メソッド</span><span class="sxs-lookup"><span data-stu-id="ff5ed-120">Method</span></span> |
| [<span data-ttu-id="ff5ed-121">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="ff5ed-121">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="ff5ed-122">メソッド</span><span class="sxs-lookup"><span data-stu-id="ff5ed-122">Method</span></span> |
| [<span data-ttu-id="ff5ed-123">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="ff5ed-123">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="ff5ed-124">メソッド</span><span class="sxs-lookup"><span data-stu-id="ff5ed-124">Method</span></span> |
| [<span data-ttu-id="ff5ed-125">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="ff5ed-125">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="ff5ed-126">メソッド</span><span class="sxs-lookup"><span data-stu-id="ff5ed-126">Method</span></span> |
| [<span data-ttu-id="ff5ed-127">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="ff5ed-127">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="ff5ed-128">メソッド</span><span class="sxs-lookup"><span data-stu-id="ff5ed-128">Method</span></span> |
| [<span data-ttu-id="ff5ed-129">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="ff5ed-129">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="ff5ed-130">メソッド</span><span class="sxs-lookup"><span data-stu-id="ff5ed-130">Method</span></span> |
| [<span data-ttu-id="ff5ed-131">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="ff5ed-131">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="ff5ed-132">メソッド</span><span class="sxs-lookup"><span data-stu-id="ff5ed-132">Method</span></span> |
| [<span data-ttu-id="ff5ed-133">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="ff5ed-133">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="ff5ed-134">メソッド</span><span class="sxs-lookup"><span data-stu-id="ff5ed-134">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="ff5ed-135">名前空間</span><span class="sxs-lookup"><span data-stu-id="ff5ed-135">Namespaces</span></span>

<span data-ttu-id="ff5ed-136">[diagnostics](Office.context.mailbox.diagnostics.md):Outlook アドインに診断情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-136">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="ff5ed-137">[item](Office.context.mailbox.item.md):Outlook アドインのメッセージや予定にアクセスするためのメソッドとプロパティを提供します。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-137">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="ff5ed-138">[userProfile](Office.context.mailbox.userProfile.md):Outlook アドインのユーザーに関する情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-138">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="ff5ed-139">メンバー</span><span class="sxs-lookup"><span data-stu-id="ff5ed-139">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="ff5ed-140">ewsUrl: String</span><span class="sxs-lookup"><span data-stu-id="ff5ed-140">ewsUrl: String</span></span>

<span data-ttu-id="ff5ed-141">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account.</span><span class="sxs-lookup"><span data-stu-id="ff5ed-141">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account.</span></span> <span data-ttu-id="ff5ed-142">Read mode only.</span><span class="sxs-lookup"><span data-stu-id="ff5ed-142">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="ff5ed-143">このメンバーは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-143">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="ff5ed-p102">
  `ewsUrl\` 値は、リモート サービスで、ユーザーのメールボックスに EWS 呼び出しを行うために使うことができます。たとえば、[選択したアイテムから添付ファイルを取得する](/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

##### <a name="type"></a><span data-ttu-id="ff5ed-146">型</span><span class="sxs-lookup"><span data-stu-id="ff5ed-146">Type</span></span>

*   <span data-ttu-id="ff5ed-147">String</span><span class="sxs-lookup"><span data-stu-id="ff5ed-147">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ff5ed-148">要件</span><span class="sxs-lookup"><span data-stu-id="ff5ed-148">Requirements</span></span>

|<span data-ttu-id="ff5ed-149">要件</span><span class="sxs-lookup"><span data-stu-id="ff5ed-149">Requirement</span></span>| <span data-ttu-id="ff5ed-150">値</span><span class="sxs-lookup"><span data-stu-id="ff5ed-150">Value</span></span>|
|---|---|
|[<span data-ttu-id="ff5ed-151">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ff5ed-151">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ff5ed-152">1.0</span><span class="sxs-lookup"><span data-stu-id="ff5ed-152">1.0</span></span>|
|[<span data-ttu-id="ff5ed-153">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ff5ed-153">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ff5ed-154">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ff5ed-154">ReadItem</span></span>|
|[<span data-ttu-id="ff5ed-155">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ff5ed-155">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ff5ed-156">読み取り</span><span class="sxs-lookup"><span data-stu-id="ff5ed-156">Read</span></span>|

### <a name="methods"></a><span data-ttu-id="ff5ed-157">メソッド</span><span class="sxs-lookup"><span data-stu-id="ff5ed-157">Methods</span></span>

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-12"></a><span data-ttu-id="ff5ed-158">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.2)}</span><span class="sxs-lookup"><span data-stu-id="ff5ed-158">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.2)}</span></span>

<span data-ttu-id="ff5ed-159">クライアントのローカル時間で時間情報が含まれている辞書を取得します。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-159">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="ff5ed-160">デスクトップまたは web 上の Outlook 用メールアプリは、日付と時刻に異なるタイムゾーンを使用できます。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-160">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times.</span></span> <span data-ttu-id="ff5ed-161">デスクトップ上の Outlook では、クライアントコンピューターのタイムゾーンが使用されます。Outlook on the web では、Exchange 管理センター (EAC) で設定されているタイムゾーンが使用されます。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-161">Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC).</span></span> <span data-ttu-id="ff5ed-162">日付と時刻の値を処理して、ユーザーインターフェイスに表示される値が、ユーザーが期待するタイムゾーンに常に一致するようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-162">You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="ff5ed-163">デスクトップクライアント上の Outlook でメールアプリが実行されている`convertToLocalClientTime`場合、このメソッドは、クライアントコンピューターのタイムゾーンに設定された値を持つ dictionary オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-163">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone.</span></span> <span data-ttu-id="ff5ed-164">メールアプリが web 上の Outlook で実行されている`convertToLocalClientTime`場合、このメソッドは、EAC で指定されたタイムゾーンに設定された値を持つ dictionary オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-164">If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ff5ed-165">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ff5ed-165">Parameters</span></span>

|<span data-ttu-id="ff5ed-166">名前</span><span class="sxs-lookup"><span data-stu-id="ff5ed-166">Name</span></span>| <span data-ttu-id="ff5ed-167">種類</span><span class="sxs-lookup"><span data-stu-id="ff5ed-167">Type</span></span>| <span data-ttu-id="ff5ed-168">説明</span><span class="sxs-lookup"><span data-stu-id="ff5ed-168">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="ff5ed-169">日付</span><span class="sxs-lookup"><span data-stu-id="ff5ed-169">Date</span></span>|<span data-ttu-id="ff5ed-170">日付オブジェクト</span><span class="sxs-lookup"><span data-stu-id="ff5ed-170">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ff5ed-171">要件</span><span class="sxs-lookup"><span data-stu-id="ff5ed-171">Requirements</span></span>

|<span data-ttu-id="ff5ed-172">要件</span><span class="sxs-lookup"><span data-stu-id="ff5ed-172">Requirement</span></span>| <span data-ttu-id="ff5ed-173">値</span><span class="sxs-lookup"><span data-stu-id="ff5ed-173">Value</span></span>|
|---|---|
|[<span data-ttu-id="ff5ed-174">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ff5ed-174">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ff5ed-175">1.0</span><span class="sxs-lookup"><span data-stu-id="ff5ed-175">1.0</span></span>|
|[<span data-ttu-id="ff5ed-176">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ff5ed-176">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ff5ed-177">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ff5ed-177">ReadItem</span></span>|
|[<span data-ttu-id="ff5ed-178">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ff5ed-178">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ff5ed-179">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ff5ed-179">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ff5ed-180">戻り値:</span><span class="sxs-lookup"><span data-stu-id="ff5ed-180">Returns:</span></span>

<span data-ttu-id="ff5ed-181">型:[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="ff5ed-181">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.2)</span></span>

<br>

---
---

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="ff5ed-182">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="ff5ed-182">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="ff5ed-183">時間情報が含まれているディクショナリから日付オブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-183">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="ff5ed-184">`convertToUtcClientTime` メソッドは、ローカルの日付と時刻を含むディクショナリを、ローカルの日付と時刻の正しい値を持つ日付オブジェクトに変換します。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-184">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ff5ed-185">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ff5ed-185">Parameters</span></span>

|<span data-ttu-id="ff5ed-186">名前</span><span class="sxs-lookup"><span data-stu-id="ff5ed-186">Name</span></span>| <span data-ttu-id="ff5ed-187">種類</span><span class="sxs-lookup"><span data-stu-id="ff5ed-187">Type</span></span>| <span data-ttu-id="ff5ed-188">説明</span><span class="sxs-lookup"><span data-stu-id="ff5ed-188">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="ff5ed-189">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="ff5ed-189">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.2)|<span data-ttu-id="ff5ed-190">変換するローカル時刻の値。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-190">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ff5ed-191">要件</span><span class="sxs-lookup"><span data-stu-id="ff5ed-191">Requirements</span></span>

|<span data-ttu-id="ff5ed-192">要件</span><span class="sxs-lookup"><span data-stu-id="ff5ed-192">Requirement</span></span>| <span data-ttu-id="ff5ed-193">値</span><span class="sxs-lookup"><span data-stu-id="ff5ed-193">Value</span></span>|
|---|---|
|[<span data-ttu-id="ff5ed-194">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ff5ed-194">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ff5ed-195">1.0</span><span class="sxs-lookup"><span data-stu-id="ff5ed-195">1.0</span></span>|
|[<span data-ttu-id="ff5ed-196">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ff5ed-196">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ff5ed-197">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ff5ed-197">ReadItem</span></span>|
|[<span data-ttu-id="ff5ed-198">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ff5ed-198">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ff5ed-199">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ff5ed-199">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ff5ed-200">戻り値:</span><span class="sxs-lookup"><span data-stu-id="ff5ed-200">Returns:</span></span>

<span data-ttu-id="ff5ed-201">時間が UTC で表現された日付オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-201">A Date object with the time expressed in UTC.</span></span>

<span data-ttu-id="ff5ed-202">型: Date</span><span class="sxs-lookup"><span data-stu-id="ff5ed-202">Type: Date</span></span>

##### <a name="example"></a><span data-ttu-id="ff5ed-203">例</span><span class="sxs-lookup"><span data-stu-id="ff5ed-203">Example</span></span>

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

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="ff5ed-204">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="ff5ed-204">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="ff5ed-205">既存の予定を表示します。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-205">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="ff5ed-206">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-206">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="ff5ed-207">`displayAppointmentForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存の予定を開きます。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-207">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="ff5ed-208">Outlook on the Mac では、このメソッドを使用して、定期的なアイテムの一部ではない単一の予定を表示したり、定期的なアイテムのマスター予定を表示したりすることはできませんが、一連のインスタンスを表示することはできません。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-208">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series.</span></span> <span data-ttu-id="ff5ed-209">これは、Mac 上の Outlook では、定期的なアイテムのインスタンスのプロパティ (アイテム ID を含む) にアクセスできないためです。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-209">This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="ff5ed-210">Web 上の Outlook では、フォームの本文が 32 KB 以下の文字である場合にのみ、このメソッドは指定されたフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-210">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="ff5ed-211">指定のアイテム識別子が既存の予定を表していない場合は、クライアント コンピューターまたはデバイスで空のウィンドウが開き、エラー メッセージは返されません。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-211">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ff5ed-212">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ff5ed-212">Parameters</span></span>

|<span data-ttu-id="ff5ed-213">名前</span><span class="sxs-lookup"><span data-stu-id="ff5ed-213">Name</span></span>| <span data-ttu-id="ff5ed-214">種類</span><span class="sxs-lookup"><span data-stu-id="ff5ed-214">Type</span></span>| <span data-ttu-id="ff5ed-215">説明</span><span class="sxs-lookup"><span data-stu-id="ff5ed-215">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="ff5ed-216">String</span><span class="sxs-lookup"><span data-stu-id="ff5ed-216">String</span></span>|<span data-ttu-id="ff5ed-217">既存の予定の Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-217">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ff5ed-218">要件</span><span class="sxs-lookup"><span data-stu-id="ff5ed-218">Requirements</span></span>

|<span data-ttu-id="ff5ed-219">要件</span><span class="sxs-lookup"><span data-stu-id="ff5ed-219">Requirement</span></span>| <span data-ttu-id="ff5ed-220">値</span><span class="sxs-lookup"><span data-stu-id="ff5ed-220">Value</span></span>|
|---|---|
|[<span data-ttu-id="ff5ed-221">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ff5ed-221">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ff5ed-222">1.0</span><span class="sxs-lookup"><span data-stu-id="ff5ed-222">1.0</span></span>|
|[<span data-ttu-id="ff5ed-223">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ff5ed-223">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ff5ed-224">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ff5ed-224">ReadItem</span></span>|
|[<span data-ttu-id="ff5ed-225">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ff5ed-225">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ff5ed-226">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ff5ed-226">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ff5ed-227">例</span><span class="sxs-lookup"><span data-stu-id="ff5ed-227">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

<br>

---
---

#### <a name="displaymessageformitemid"></a><span data-ttu-id="ff5ed-228">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="ff5ed-228">displayMessageForm(itemId)</span></span>

<span data-ttu-id="ff5ed-229">既存のメッセージを表示します。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-229">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="ff5ed-230">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-230">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="ff5ed-231">`displayMessageForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存のメッセージを開きます。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-231">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="ff5ed-232">Web 上の Outlook では、フォームの本文が 32 KB の文字数以下の場合にのみ、このメソッドは指定されたフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-232">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="ff5ed-233">指定のアイテム識別子が既存のメッセージを表していない場合は、クライアント コンピューターにはメッセージは表示されず、エラー メッセージも返されません。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-233">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="ff5ed-p106">予定を表す `itemId` を含む `displayMessageForm` を使用しないでください。既存の予定を表示するには、`displayAppointmentForm` メソッドを使用します。新しい予定を作成するフォームを表示するには、`displayNewAppointmentForm` メソッドを使用します。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-p106">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ff5ed-236">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ff5ed-236">Parameters</span></span>

|<span data-ttu-id="ff5ed-237">名前</span><span class="sxs-lookup"><span data-stu-id="ff5ed-237">Name</span></span>| <span data-ttu-id="ff5ed-238">型</span><span class="sxs-lookup"><span data-stu-id="ff5ed-238">Type</span></span>| <span data-ttu-id="ff5ed-239">説明</span><span class="sxs-lookup"><span data-stu-id="ff5ed-239">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="ff5ed-240">String</span><span class="sxs-lookup"><span data-stu-id="ff5ed-240">String</span></span>|<span data-ttu-id="ff5ed-241">既存のメッセージの Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-241">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ff5ed-242">要件</span><span class="sxs-lookup"><span data-stu-id="ff5ed-242">Requirements</span></span>

|<span data-ttu-id="ff5ed-243">要件</span><span class="sxs-lookup"><span data-stu-id="ff5ed-243">Requirement</span></span>| <span data-ttu-id="ff5ed-244">値</span><span class="sxs-lookup"><span data-stu-id="ff5ed-244">Value</span></span>|
|---|---|
|[<span data-ttu-id="ff5ed-245">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ff5ed-245">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ff5ed-246">1.0</span><span class="sxs-lookup"><span data-stu-id="ff5ed-246">1.0</span></span>|
|[<span data-ttu-id="ff5ed-247">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ff5ed-247">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ff5ed-248">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ff5ed-248">ReadItem</span></span>|
|[<span data-ttu-id="ff5ed-249">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ff5ed-249">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ff5ed-250">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ff5ed-250">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ff5ed-251">例</span><span class="sxs-lookup"><span data-stu-id="ff5ed-251">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

<br>

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="ff5ed-252">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="ff5ed-252">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="ff5ed-253">新しい予定を作成するためのフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-253">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="ff5ed-254">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-254">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="ff5ed-p107">`displayNewAppointmentForm` メソッドを使用すると、ユーザーが新しい予定または会議を作成できるフォームが開きます。パラメーターを指定すると、予定のフォーム フィールドにパラメーターの内容が自動的に設定されます。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-p107">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="ff5ed-257">Outlook on the web およびモバイルデバイスでは、このメソッドは常に出席者フィールドを含むフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-257">In Outlook on the web and mobile devices, this method always displays a form with an attendees field.</span></span> <span data-ttu-id="ff5ed-258">入力引数として出席者を指定しないと、このメソッドにより **[保存]** ボタンのあるフォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-258">If you do not specify any attendees as input arguments, the method displays a form with a **Save** button.</span></span> <span data-ttu-id="ff5ed-259">出席者を指定した場合には、フォームにその出席者と **[送信]** ボタンが表示されます。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-259">If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="ff5ed-p109">Outlook リッチ クライアントと Outlook RT で、`requiredAttendees`、`optionalAttendees`、または `resources` パラメーターに出席者またはリソースを指定し、このメソッドを実行すると、**[送信]** ボタンがある会議フォームが表示されます。受信者を指定せずにこのメソッドを実行すると、**[保存して閉じる]** ボタンがある予定フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-p109">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="ff5ed-262">パラメーターのいずれかが指定のサイズ制限を超える場合、または不明なパラメーター名が指定されている場合は、例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-262">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ff5ed-263">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ff5ed-263">Parameters</span></span>

|<span data-ttu-id="ff5ed-264">名前</span><span class="sxs-lookup"><span data-stu-id="ff5ed-264">Name</span></span>| <span data-ttu-id="ff5ed-265">型</span><span class="sxs-lookup"><span data-stu-id="ff5ed-265">Type</span></span>| <span data-ttu-id="ff5ed-266">説明</span><span class="sxs-lookup"><span data-stu-id="ff5ed-266">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="ff5ed-267">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="ff5ed-267">Object</span></span> | <span data-ttu-id="ff5ed-268">新しい予定を記述するパラメーターのディクショナリ。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-268">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="ff5ed-269">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)&gt;</span><span class="sxs-lookup"><span data-stu-id="ff5ed-269">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)&gt;</span></span> | <span data-ttu-id="ff5ed-p110">予定に必要な各出席者について、メール アドレスを含む文字列の配列、または `EmailAddressDetails` オブジェクトを含む配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-p110">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="ff5ed-272">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)&gt;</span><span class="sxs-lookup"><span data-stu-id="ff5ed-272">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)&gt;</span></span> | <span data-ttu-id="ff5ed-p111">予定の各任意出席者について、メール アドレスを含む文字列の配列、または `EmailAddressDetails` オブジェクトを含む配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-p111">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="ff5ed-275">日付</span><span class="sxs-lookup"><span data-stu-id="ff5ed-275">Date</span></span> | <span data-ttu-id="ff5ed-276">予定の開始日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-276">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="ff5ed-277">日付</span><span class="sxs-lookup"><span data-stu-id="ff5ed-277">Date</span></span> | <span data-ttu-id="ff5ed-278">予定の終了日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-278">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="ff5ed-279">String</span><span class="sxs-lookup"><span data-stu-id="ff5ed-279">String</span></span> | <span data-ttu-id="ff5ed-p112">予定の場所を含む文字列。文字列は最大 255 文字に制限されます。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-p112">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="ff5ed-282">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="ff5ed-282">Array.&lt;String&gt;</span></span> | <span data-ttu-id="ff5ed-p113">予定に必要なリソースを含む文字列の配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-p113">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="ff5ed-285">String</span><span class="sxs-lookup"><span data-stu-id="ff5ed-285">String</span></span> | <span data-ttu-id="ff5ed-p114">予定の件名を含む文字列です。文字列は最大 255 文字に制限されます。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-p114">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="ff5ed-288">String</span><span class="sxs-lookup"><span data-stu-id="ff5ed-288">String</span></span> | <span data-ttu-id="ff5ed-p115">予定の本文。本文の内容は、最大サイズが 32 KB に制限されます。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-p115">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="ff5ed-291">要件</span><span class="sxs-lookup"><span data-stu-id="ff5ed-291">Requirements</span></span>

|<span data-ttu-id="ff5ed-292">要件</span><span class="sxs-lookup"><span data-stu-id="ff5ed-292">Requirement</span></span>| <span data-ttu-id="ff5ed-293">値</span><span class="sxs-lookup"><span data-stu-id="ff5ed-293">Value</span></span>|
|---|---|
|[<span data-ttu-id="ff5ed-294">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ff5ed-294">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ff5ed-295">1.0</span><span class="sxs-lookup"><span data-stu-id="ff5ed-295">1.0</span></span>|
|[<span data-ttu-id="ff5ed-296">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ff5ed-296">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ff5ed-297">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ff5ed-297">ReadItem</span></span>|
|[<span data-ttu-id="ff5ed-298">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ff5ed-298">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ff5ed-299">読み取り</span><span class="sxs-lookup"><span data-stu-id="ff5ed-299">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ff5ed-300">例</span><span class="sxs-lookup"><span data-stu-id="ff5ed-300">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="ff5ed-301">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="ff5ed-301">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="ff5ed-302">Exchange Server から添付ファイルやアイテムを取得するために使うトークンを含む文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-302">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="ff5ed-p116">`getCallbackTokenAsync` メソッドは、ユーザーのメールボックスをホストする Exchange Server から不透明なトークンを取得する非同期の呼び出しを行います。コールバック トークンの有効期間は 5 分です。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-p116">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="ff5ed-p117">トークンと予定の識別子またはアイテムの識別子をサードパーティ システムに渡すことができます。サードパーティ システムは、トークンをベアラー承認トークンとして使用し、Exchange Web サービス (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) または [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) 操作を呼び出して、添付ファイルまたはアイテムを返します。たとえば、[選択したアイテムから添付ファイルを取得する](/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-p117">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="ff5ed-308">アプリが \*\*\*\* メソッドを呼び出すには、アプリのマニフェスト内に `getCallbackTokenAsync` アクセス許可が指定されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-308">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ff5ed-309">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ff5ed-309">Parameters</span></span>

|<span data-ttu-id="ff5ed-310">名前</span><span class="sxs-lookup"><span data-stu-id="ff5ed-310">Name</span></span>| <span data-ttu-id="ff5ed-311">型</span><span class="sxs-lookup"><span data-stu-id="ff5ed-311">Type</span></span>| <span data-ttu-id="ff5ed-312">属性</span><span class="sxs-lookup"><span data-stu-id="ff5ed-312">Attributes</span></span>| <span data-ttu-id="ff5ed-313">説明</span><span class="sxs-lookup"><span data-stu-id="ff5ed-313">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="ff5ed-314">function</span><span class="sxs-lookup"><span data-stu-id="ff5ed-314">function</span></span>||<span data-ttu-id="ff5ed-315">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-315">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="ff5ed-316">トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-316">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="ff5ed-317">エラーが発生した場合は`asyncResult.error` 、 `asyncResult.diagnostics`プロパティとプロパティによって追加情報が提供されることがあります。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-317">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="ff5ed-318">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="ff5ed-318">Object</span></span>| <span data-ttu-id="ff5ed-319">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="ff5ed-319">&lt;optional&gt;</span></span>|<span data-ttu-id="ff5ed-320">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-320">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="ff5ed-321">エラー</span><span class="sxs-lookup"><span data-stu-id="ff5ed-321">Errors</span></span>

|<span data-ttu-id="ff5ed-322">エラー コード</span><span class="sxs-lookup"><span data-stu-id="ff5ed-322">Error code</span></span>|<span data-ttu-id="ff5ed-323">説明</span><span class="sxs-lookup"><span data-stu-id="ff5ed-323">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="ff5ed-324">要求が失敗しました。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-324">The request has failed.</span></span> <span data-ttu-id="ff5ed-325">HTTP エラーコードについては、diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-325">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="ff5ed-326">Exchange サーバーがエラーを返しました。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-326">The Exchange server returned an error.</span></span> <span data-ttu-id="ff5ed-327">詳細については、「diagnostics オブジェクト」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-327">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="ff5ed-328">ユーザーがネットワークに接続されていません。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-328">The user is no longer connected to the network.</span></span> <span data-ttu-id="ff5ed-329">ネットワーク接続を確認し、もう一度実行してください。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-329">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ff5ed-330">要件</span><span class="sxs-lookup"><span data-stu-id="ff5ed-330">Requirements</span></span>

|<span data-ttu-id="ff5ed-331">要件</span><span class="sxs-lookup"><span data-stu-id="ff5ed-331">Requirement</span></span>| <span data-ttu-id="ff5ed-332">値</span><span class="sxs-lookup"><span data-stu-id="ff5ed-332">Value</span></span>|
|---|---|
|[<span data-ttu-id="ff5ed-333">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ff5ed-333">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ff5ed-334">1.0</span><span class="sxs-lookup"><span data-stu-id="ff5ed-334">1.0</span></span>|
|[<span data-ttu-id="ff5ed-335">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ff5ed-335">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ff5ed-336">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ff5ed-336">ReadItem</span></span>|
|[<span data-ttu-id="ff5ed-337">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ff5ed-337">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ff5ed-338">読み取り</span><span class="sxs-lookup"><span data-stu-id="ff5ed-338">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ff5ed-339">例</span><span class="sxs-lookup"><span data-stu-id="ff5ed-339">Example</span></span>

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

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="ff5ed-340">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="ff5ed-340">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="ff5ed-341">ユーザーと Office アドインを識別するトークンを取得します。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-341">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="ff5ed-342">`getUserIdentityTokenAsync` メソッドは、[アドインとユーザーをサード パーティのシステムで識別して認証](/outlook/add-ins/authentication)することのできるトークンを返します。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-342">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="ff5ed-343">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ff5ed-343">Parameters</span></span>

|<span data-ttu-id="ff5ed-344">名前</span><span class="sxs-lookup"><span data-stu-id="ff5ed-344">Name</span></span>| <span data-ttu-id="ff5ed-345">型</span><span class="sxs-lookup"><span data-stu-id="ff5ed-345">Type</span></span>| <span data-ttu-id="ff5ed-346">属性</span><span class="sxs-lookup"><span data-stu-id="ff5ed-346">Attributes</span></span>| <span data-ttu-id="ff5ed-347">説明</span><span class="sxs-lookup"><span data-stu-id="ff5ed-347">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="ff5ed-348">function</span><span class="sxs-lookup"><span data-stu-id="ff5ed-348">function</span></span>||<span data-ttu-id="ff5ed-349">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-349">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="ff5ed-350">トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-350">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="ff5ed-351">エラーが発生した場合は`asyncResult.error` 、 `asyncResult.diagnostics`プロパティとプロパティによって追加情報が提供されることがあります。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-351">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="ff5ed-352">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="ff5ed-352">Object</span></span>| <span data-ttu-id="ff5ed-353">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="ff5ed-353">&lt;optional&gt;</span></span>|<span data-ttu-id="ff5ed-354">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-354">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="ff5ed-355">エラー</span><span class="sxs-lookup"><span data-stu-id="ff5ed-355">Errors</span></span>

|<span data-ttu-id="ff5ed-356">エラー コード</span><span class="sxs-lookup"><span data-stu-id="ff5ed-356">Error code</span></span>|<span data-ttu-id="ff5ed-357">説明</span><span class="sxs-lookup"><span data-stu-id="ff5ed-357">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="ff5ed-358">要求が失敗しました。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-358">The request has failed.</span></span> <span data-ttu-id="ff5ed-359">HTTP エラーコードについては、diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-359">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="ff5ed-360">Exchange サーバーがエラーを返しました。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-360">The Exchange server returned an error.</span></span> <span data-ttu-id="ff5ed-361">詳細については、「diagnostics オブジェクト」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-361">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="ff5ed-362">ユーザーがネットワークに接続されていません。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-362">The user is no longer connected to the network.</span></span> <span data-ttu-id="ff5ed-363">ネットワーク接続を確認し、もう一度実行してください。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-363">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ff5ed-364">要件</span><span class="sxs-lookup"><span data-stu-id="ff5ed-364">Requirements</span></span>

|<span data-ttu-id="ff5ed-365">要件</span><span class="sxs-lookup"><span data-stu-id="ff5ed-365">Requirement</span></span>| <span data-ttu-id="ff5ed-366">値</span><span class="sxs-lookup"><span data-stu-id="ff5ed-366">Value</span></span>|
|---|---|
|[<span data-ttu-id="ff5ed-367">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ff5ed-367">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ff5ed-368">1.0</span><span class="sxs-lookup"><span data-stu-id="ff5ed-368">1.0</span></span>|
|[<span data-ttu-id="ff5ed-369">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ff5ed-369">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ff5ed-370">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ff5ed-370">ReadItem</span></span>|
|[<span data-ttu-id="ff5ed-371">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ff5ed-371">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ff5ed-372">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ff5ed-372">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ff5ed-373">例</span><span class="sxs-lookup"><span data-stu-id="ff5ed-373">Example</span></span>

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

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="ff5ed-374">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="ff5ed-374">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="ff5ed-375">ユーザーのメールボックスをホストしている Exchange サーバー上の Exchange Web サービス (EWS) のサービスに対して非同期の要求を行います。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-375">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="ff5ed-376">このメソッドは、次のシナリオではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-376">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="ff5ed-377">Outlook on iOS または Android</span><span class="sxs-lookup"><span data-stu-id="ff5ed-377">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="ff5ed-378">アドインが Gmail のメールボックスに読み込まれる場合</span><span class="sxs-lookup"><span data-stu-id="ff5ed-378">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="ff5ed-379">このような場合は、アドインでは [REST API を使用](/outlook/add-ins/use-rest-api)して、代わりにユーザーのメールボックスにアクセスする必要があります。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-379">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="ff5ed-p124">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange. See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span><span class="sxs-lookup"><span data-stu-id="ff5ed-p124">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange. See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="ff5ed-382">`makeEwsRequestAsync` メソッドでは、フォルダー関連アイテムを要求できません。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-382">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="ff5ed-383">XML 要求では UTF-8 エンコードを指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-383">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="ff5ed-p125">`makeEwsRequestAsync` メソッドを使用するには、アドインに **ReadWriteMailbox** アクセス許可が必要です。**ReadWriteMailbox** アクセス許可と、`makeEwsRequestAsync` メソッドで呼び出せる EWS 操作の使い方については、「[ユーザーのメールボックスへのメール アドイン アクセスのアクセス許可を指定する](/outlook/add-ins/understanding-outlook-add-in-permissions)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-p125">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="ff5ed-386">サーバー管理者は、クライアント アクセス サーバーの EWS ディレクトリで `OAuthAuthentication` を true に設定して、`makeEwsRequestAsync` メソッドで EWS 要求を行うことができるようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-386">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="ff5ed-387">バージョンの相違点</span><span class="sxs-lookup"><span data-stu-id="ff5ed-387">Version differences</span></span>

<span data-ttu-id="ff5ed-388">バージョン 15.0.4535.1004 より前のバージョンの Outlook で実行しているメール アプリで `makeEwsRequestAsync` メソッドを使う場合は、エンコード値を `ISO-8859-1` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-388">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="ff5ed-p126">Outlook on the web でメール アプリを実行している場合は、エンコード値を設定する必要はありません。mailbox.diagnostics.hostName プロパティを使って、メール アプリを Outlook で実行しているのか、Outlook on the web で実行しているのかを確認できます。mailbox.diagnostics.hostVersion プロパティを使って、どのバージョンの Outlook を使って実行しているかを確認できます。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-p126">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ff5ed-392">パラメーター</span><span class="sxs-lookup"><span data-stu-id="ff5ed-392">Parameters</span></span>

|<span data-ttu-id="ff5ed-393">名前</span><span class="sxs-lookup"><span data-stu-id="ff5ed-393">Name</span></span>| <span data-ttu-id="ff5ed-394">型</span><span class="sxs-lookup"><span data-stu-id="ff5ed-394">Type</span></span>| <span data-ttu-id="ff5ed-395">属性</span><span class="sxs-lookup"><span data-stu-id="ff5ed-395">Attributes</span></span>| <span data-ttu-id="ff5ed-396">説明</span><span class="sxs-lookup"><span data-stu-id="ff5ed-396">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="ff5ed-397">String</span><span class="sxs-lookup"><span data-stu-id="ff5ed-397">String</span></span>||<span data-ttu-id="ff5ed-398">EWS 要求です。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-398">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="ff5ed-399">function</span><span class="sxs-lookup"><span data-stu-id="ff5ed-399">function</span></span>||<span data-ttu-id="ff5ed-400">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-400">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="ff5ed-p127">The XML result of the EWS call is provided as a string in the `asyncResult.value` property. If the result exceeds 1 MB in size, an error message is returned instead.</span><span class="sxs-lookup"><span data-stu-id="ff5ed-p127">The XML result of the EWS call is provided as a string in the `asyncResult.value` property. If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="ff5ed-403">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="ff5ed-403">Object</span></span>| <span data-ttu-id="ff5ed-404">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="ff5ed-404">&lt;optional&gt;</span></span>|<span data-ttu-id="ff5ed-405">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-405">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ff5ed-406">要件</span><span class="sxs-lookup"><span data-stu-id="ff5ed-406">Requirements</span></span>

|<span data-ttu-id="ff5ed-407">要件</span><span class="sxs-lookup"><span data-stu-id="ff5ed-407">Requirement</span></span>| <span data-ttu-id="ff5ed-408">値</span><span class="sxs-lookup"><span data-stu-id="ff5ed-408">Value</span></span>|
|---|---|
|[<span data-ttu-id="ff5ed-409">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ff5ed-409">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ff5ed-410">1.0</span><span class="sxs-lookup"><span data-stu-id="ff5ed-410">1.0</span></span>|
|[<span data-ttu-id="ff5ed-411">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ff5ed-411">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ff5ed-412">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="ff5ed-412">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="ff5ed-413">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ff5ed-413">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ff5ed-414">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ff5ed-414">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ff5ed-415">例</span><span class="sxs-lookup"><span data-stu-id="ff5ed-415">Example</span></span>

<span data-ttu-id="ff5ed-416">次の例は、`makeEwsRequestAsync` を呼び出し、`GetItem` 操作を使って項目の件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="ff5ed-416">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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
