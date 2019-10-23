---
title: Office. メールボックス要件セット1.2
description: ''
ms.date: 10/21/2019
localization_priority: Normal
ms.openlocfilehash: 542e8c9899c2d4a3c5b4546c3d5a73ba0d3c3a7e
ms.sourcegitcommit: 499bf49b41205f8034c501d4db5fe4b02dab205e
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/22/2019
ms.locfileid: "37627000"
---
# <a name="mailbox"></a><span data-ttu-id="1fbec-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="1fbec-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="1fbec-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="1fbec-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="1fbec-104">Microsoft Outlook の Outlook アドイン オブジェクト モデルへのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="1fbec-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1fbec-105">要件</span><span class="sxs-lookup"><span data-stu-id="1fbec-105">Requirements</span></span>

|<span data-ttu-id="1fbec-106">要件</span><span class="sxs-lookup"><span data-stu-id="1fbec-106">Requirement</span></span>| <span data-ttu-id="1fbec-107">値</span><span class="sxs-lookup"><span data-stu-id="1fbec-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fbec-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1fbec-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fbec-109">1.0</span><span class="sxs-lookup"><span data-stu-id="1fbec-109">1.0</span></span>|
|[<span data-ttu-id="1fbec-110">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1fbec-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fbec-111">制限あり</span><span class="sxs-lookup"><span data-stu-id="1fbec-111">Restricted</span></span>|
|[<span data-ttu-id="1fbec-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1fbec-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1fbec-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1fbec-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="1fbec-114">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="1fbec-114">Members and methods</span></span>

| <span data-ttu-id="1fbec-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="1fbec-115">Member</span></span> | <span data-ttu-id="1fbec-116">種類</span><span class="sxs-lookup"><span data-stu-id="1fbec-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="1fbec-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="1fbec-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="1fbec-118">メンバー</span><span class="sxs-lookup"><span data-stu-id="1fbec-118">Member</span></span> |
| [<span data-ttu-id="1fbec-119">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="1fbec-119">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="1fbec-120">メソッド</span><span class="sxs-lookup"><span data-stu-id="1fbec-120">Method</span></span> |
| [<span data-ttu-id="1fbec-121">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="1fbec-121">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="1fbec-122">メソッド</span><span class="sxs-lookup"><span data-stu-id="1fbec-122">Method</span></span> |
| [<span data-ttu-id="1fbec-123">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="1fbec-123">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="1fbec-124">メソッド</span><span class="sxs-lookup"><span data-stu-id="1fbec-124">Method</span></span> |
| [<span data-ttu-id="1fbec-125">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="1fbec-125">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="1fbec-126">メソッド</span><span class="sxs-lookup"><span data-stu-id="1fbec-126">Method</span></span> |
| [<span data-ttu-id="1fbec-127">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="1fbec-127">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="1fbec-128">メソッド</span><span class="sxs-lookup"><span data-stu-id="1fbec-128">Method</span></span> |
| [<span data-ttu-id="1fbec-129">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="1fbec-129">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="1fbec-130">メソッド</span><span class="sxs-lookup"><span data-stu-id="1fbec-130">Method</span></span> |
| [<span data-ttu-id="1fbec-131">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="1fbec-131">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="1fbec-132">メソッド</span><span class="sxs-lookup"><span data-stu-id="1fbec-132">Method</span></span> |
| [<span data-ttu-id="1fbec-133">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="1fbec-133">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="1fbec-134">メソッド</span><span class="sxs-lookup"><span data-stu-id="1fbec-134">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="1fbec-135">名前空間</span><span class="sxs-lookup"><span data-stu-id="1fbec-135">Namespaces</span></span>

<span data-ttu-id="1fbec-136">[diagnostics](Office.context.mailbox.diagnostics.md):Outlook アドインに診断情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="1fbec-136">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="1fbec-137">[item](Office.context.mailbox.item.md):Outlook アドインのメッセージや予定にアクセスするためのメソッドとプロパティを提供します。</span><span class="sxs-lookup"><span data-stu-id="1fbec-137">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="1fbec-138">[userProfile](Office.context.mailbox.userProfile.md):Outlook アドインのユーザーに関する情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="1fbec-138">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="1fbec-139">Members</span><span class="sxs-lookup"><span data-stu-id="1fbec-139">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="1fbec-140">ewsUrl: String</span><span class="sxs-lookup"><span data-stu-id="1fbec-140">ewsUrl: String</span></span>

<span data-ttu-id="1fbec-p101">このメール アカウントの Exchange Web サービス (EWS) エンドポイントの URL を取得します。読み取りモードのみです。</span><span class="sxs-lookup"><span data-stu-id="1fbec-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="1fbec-143">このメンバーは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="1fbec-143">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="1fbec-p102">
  `ewsUrl\` 値は、リモート サービスで、ユーザーのメールボックスに EWS 呼び出しを行うために使うことができます。たとえば、[選択したアイテムから添付ファイルを取得する](/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="1fbec-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

##### <a name="type"></a><span data-ttu-id="1fbec-146">型</span><span class="sxs-lookup"><span data-stu-id="1fbec-146">Type</span></span>

*   <span data-ttu-id="1fbec-147">String</span><span class="sxs-lookup"><span data-stu-id="1fbec-147">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1fbec-148">要件</span><span class="sxs-lookup"><span data-stu-id="1fbec-148">Requirements</span></span>

|<span data-ttu-id="1fbec-149">要件</span><span class="sxs-lookup"><span data-stu-id="1fbec-149">Requirement</span></span>| <span data-ttu-id="1fbec-150">値</span><span class="sxs-lookup"><span data-stu-id="1fbec-150">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fbec-151">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1fbec-151">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fbec-152">1.0</span><span class="sxs-lookup"><span data-stu-id="1fbec-152">1.0</span></span>|
|[<span data-ttu-id="1fbec-153">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1fbec-153">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fbec-154">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1fbec-154">ReadItem</span></span>|
|[<span data-ttu-id="1fbec-155">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1fbec-155">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1fbec-156">読み取り</span><span class="sxs-lookup"><span data-stu-id="1fbec-156">Read</span></span>|

### <a name="methods"></a><span data-ttu-id="1fbec-157">メソッド</span><span class="sxs-lookup"><span data-stu-id="1fbec-157">Methods</span></span>

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-12"></a><span data-ttu-id="1fbec-158">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.2)}</span><span class="sxs-lookup"><span data-stu-id="1fbec-158">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.2)}</span></span>

<span data-ttu-id="1fbec-159">クライアントのローカル時間で時間情報が含まれている辞書を取得します。</span><span class="sxs-lookup"><span data-stu-id="1fbec-159">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="1fbec-p103">Outlook on the web または Outlook デスクトップのメールアプリでは、日付と時刻に異なるタイムゾーンを使用できます。Outlook デスクトップは、クライアント コンピューターのタイム ゾーンを使用します。Outlook on the web は、Exchange 管理センター (EAC) で設定されたタイム ゾーンを使用します。日付と時刻の値は、ユーザー インターフェイスに表示される値が、常にユーザーが期待するタイム ゾーンと一致するように処理する必要があります。</span><span class="sxs-lookup"><span data-stu-id="1fbec-p103">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times. Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="1fbec-p104">Outlook デスクトップ クライアントでメール アプリを実行している場合、`convertToLocalClientTime` メソッドは、クライアント コンピューターのタイム ゾーンに設定された値のディクショナリ オブジェクトを返します。Outlook on the web でメール アプリを実行している場合、`convertToLocalClientTime` メソッドは、EAC で指定したタイム ゾーンに設定された値のディクショナリ オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="1fbec-p104">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1fbec-165">パラメーター</span><span class="sxs-lookup"><span data-stu-id="1fbec-165">Parameters</span></span>

|<span data-ttu-id="1fbec-166">名前</span><span class="sxs-lookup"><span data-stu-id="1fbec-166">Name</span></span>| <span data-ttu-id="1fbec-167">種類</span><span class="sxs-lookup"><span data-stu-id="1fbec-167">Type</span></span>| <span data-ttu-id="1fbec-168">説明</span><span class="sxs-lookup"><span data-stu-id="1fbec-168">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="1fbec-169">日付</span><span class="sxs-lookup"><span data-stu-id="1fbec-169">Date</span></span>|<span data-ttu-id="1fbec-170">日付オブジェクト</span><span class="sxs-lookup"><span data-stu-id="1fbec-170">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1fbec-171">要件</span><span class="sxs-lookup"><span data-stu-id="1fbec-171">Requirements</span></span>

|<span data-ttu-id="1fbec-172">要件</span><span class="sxs-lookup"><span data-stu-id="1fbec-172">Requirement</span></span>| <span data-ttu-id="1fbec-173">値</span><span class="sxs-lookup"><span data-stu-id="1fbec-173">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fbec-174">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1fbec-174">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fbec-175">1.0</span><span class="sxs-lookup"><span data-stu-id="1fbec-175">1.0</span></span>|
|[<span data-ttu-id="1fbec-176">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1fbec-176">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fbec-177">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1fbec-177">ReadItem</span></span>|
|[<span data-ttu-id="1fbec-178">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1fbec-178">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1fbec-179">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1fbec-179">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1fbec-180">戻り値:</span><span class="sxs-lookup"><span data-stu-id="1fbec-180">Returns:</span></span>

<span data-ttu-id="1fbec-181">型:[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="1fbec-181">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.2)</span></span>

<br>

---
---

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="1fbec-182">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="1fbec-182">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="1fbec-183">時間情報が含まれているディクショナリから日付オブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="1fbec-183">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="1fbec-184">`convertToUtcClientTime` メソッドは、ローカルの日付と時刻を含むディクショナリを、ローカルの日付と時刻の正しい値を持つ日付オブジェクトに変換します。</span><span class="sxs-lookup"><span data-stu-id="1fbec-184">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1fbec-185">パラメーター</span><span class="sxs-lookup"><span data-stu-id="1fbec-185">Parameters</span></span>

|<span data-ttu-id="1fbec-186">名前</span><span class="sxs-lookup"><span data-stu-id="1fbec-186">Name</span></span>| <span data-ttu-id="1fbec-187">種類</span><span class="sxs-lookup"><span data-stu-id="1fbec-187">Type</span></span>| <span data-ttu-id="1fbec-188">説明</span><span class="sxs-lookup"><span data-stu-id="1fbec-188">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="1fbec-189">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="1fbec-189">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.2)|<span data-ttu-id="1fbec-190">変換するローカル時刻の値。</span><span class="sxs-lookup"><span data-stu-id="1fbec-190">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1fbec-191">要件</span><span class="sxs-lookup"><span data-stu-id="1fbec-191">Requirements</span></span>

|<span data-ttu-id="1fbec-192">要件</span><span class="sxs-lookup"><span data-stu-id="1fbec-192">Requirement</span></span>| <span data-ttu-id="1fbec-193">値</span><span class="sxs-lookup"><span data-stu-id="1fbec-193">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fbec-194">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1fbec-194">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fbec-195">1.0</span><span class="sxs-lookup"><span data-stu-id="1fbec-195">1.0</span></span>|
|[<span data-ttu-id="1fbec-196">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1fbec-196">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fbec-197">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1fbec-197">ReadItem</span></span>|
|[<span data-ttu-id="1fbec-198">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1fbec-198">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1fbec-199">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1fbec-199">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1fbec-200">戻り値:</span><span class="sxs-lookup"><span data-stu-id="1fbec-200">Returns:</span></span>

<span data-ttu-id="1fbec-201">時間が UTC で表現された日付オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="1fbec-201">A Date object with the time expressed in UTC.</span></span>

<span data-ttu-id="1fbec-202">型: Date</span><span class="sxs-lookup"><span data-stu-id="1fbec-202">Type: Date</span></span>

##### <a name="example"></a><span data-ttu-id="1fbec-203">例</span><span class="sxs-lookup"><span data-stu-id="1fbec-203">Example</span></span>

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

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="1fbec-204">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="1fbec-204">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="1fbec-205">既存の予定を表示します。</span><span class="sxs-lookup"><span data-stu-id="1fbec-205">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="1fbec-206">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="1fbec-206">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="1fbec-207">`displayAppointmentForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存の予定を開きます。</span><span class="sxs-lookup"><span data-stu-id="1fbec-207">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="1fbec-p105">Outlook on Mac では、このメソッドを使用して定期的な系列に含まれない単発の予定や定期的な系列のマスター予定を表示できます。ただし、系列のインスタンスは表示できません。これは、Outlook on Mac では定期的な系列のインスタンスのプロパティ (アイテム ID を含む) にアクセスできないためです。</span><span class="sxs-lookup"><span data-stu-id="1fbec-p105">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="1fbec-210">Outlook on the web では、このメソッドはフォームの本文が 32KB 以下の文字数の場合にのみ指定のフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="1fbec-210">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="1fbec-211">指定のアイテム識別子が既存の予定を表していない場合は、クライアント コンピューターまたはデバイスで空のウィンドウが開き、エラー メッセージは返されません。</span><span class="sxs-lookup"><span data-stu-id="1fbec-211">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1fbec-212">パラメーター</span><span class="sxs-lookup"><span data-stu-id="1fbec-212">Parameters</span></span>

|<span data-ttu-id="1fbec-213">名前</span><span class="sxs-lookup"><span data-stu-id="1fbec-213">Name</span></span>| <span data-ttu-id="1fbec-214">種類</span><span class="sxs-lookup"><span data-stu-id="1fbec-214">Type</span></span>| <span data-ttu-id="1fbec-215">説明</span><span class="sxs-lookup"><span data-stu-id="1fbec-215">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="1fbec-216">String</span><span class="sxs-lookup"><span data-stu-id="1fbec-216">String</span></span>|<span data-ttu-id="1fbec-217">既存の予定の Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="1fbec-217">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1fbec-218">要件</span><span class="sxs-lookup"><span data-stu-id="1fbec-218">Requirements</span></span>

|<span data-ttu-id="1fbec-219">要件</span><span class="sxs-lookup"><span data-stu-id="1fbec-219">Requirement</span></span>| <span data-ttu-id="1fbec-220">値</span><span class="sxs-lookup"><span data-stu-id="1fbec-220">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fbec-221">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1fbec-221">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fbec-222">1.0</span><span class="sxs-lookup"><span data-stu-id="1fbec-222">1.0</span></span>|
|[<span data-ttu-id="1fbec-223">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1fbec-223">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fbec-224">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1fbec-224">ReadItem</span></span>|
|[<span data-ttu-id="1fbec-225">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1fbec-225">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1fbec-226">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1fbec-226">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1fbec-227">例</span><span class="sxs-lookup"><span data-stu-id="1fbec-227">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

<br>

---
---

#### <a name="displaymessageformitemid"></a><span data-ttu-id="1fbec-228">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="1fbec-228">displayMessageForm(itemId)</span></span>

<span data-ttu-id="1fbec-229">既存のメッセージを表示します。</span><span class="sxs-lookup"><span data-stu-id="1fbec-229">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="1fbec-230">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="1fbec-230">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="1fbec-231">`displayMessageForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存のメッセージを開きます。</span><span class="sxs-lookup"><span data-stu-id="1fbec-231">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="1fbec-232">Outlook on the web では、このメソッドはフォームの本文が 32 KB 以下の文字数の場合にのみ指定のフォームを開きます。</span><span class="sxs-lookup"><span data-stu-id="1fbec-232">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="1fbec-233">指定のアイテム識別子が既存のメッセージを表していない場合は、クライアント コンピューターにはメッセージは表示されず、エラー メッセージも返されません。</span><span class="sxs-lookup"><span data-stu-id="1fbec-233">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="1fbec-p106">予定を表す `itemId` を含む `displayMessageForm` を使用しないでください。既存の予定を表示するには、`displayAppointmentForm` メソッドを使用します。新しい予定を作成するフォームを表示するには、`displayNewAppointmentForm` メソッドを使用します。</span><span class="sxs-lookup"><span data-stu-id="1fbec-p106">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1fbec-236">パラメーター</span><span class="sxs-lookup"><span data-stu-id="1fbec-236">Parameters</span></span>

|<span data-ttu-id="1fbec-237">名前</span><span class="sxs-lookup"><span data-stu-id="1fbec-237">Name</span></span>| <span data-ttu-id="1fbec-238">種類</span><span class="sxs-lookup"><span data-stu-id="1fbec-238">Type</span></span>| <span data-ttu-id="1fbec-239">Description</span><span class="sxs-lookup"><span data-stu-id="1fbec-239">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="1fbec-240">String</span><span class="sxs-lookup"><span data-stu-id="1fbec-240">String</span></span>|<span data-ttu-id="1fbec-241">既存のメッセージの Exchange Web サービス (EWS) 識別子。</span><span class="sxs-lookup"><span data-stu-id="1fbec-241">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1fbec-242">要件</span><span class="sxs-lookup"><span data-stu-id="1fbec-242">Requirements</span></span>

|<span data-ttu-id="1fbec-243">要件</span><span class="sxs-lookup"><span data-stu-id="1fbec-243">Requirement</span></span>| <span data-ttu-id="1fbec-244">値</span><span class="sxs-lookup"><span data-stu-id="1fbec-244">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fbec-245">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1fbec-245">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fbec-246">1.0</span><span class="sxs-lookup"><span data-stu-id="1fbec-246">1.0</span></span>|
|[<span data-ttu-id="1fbec-247">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1fbec-247">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fbec-248">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1fbec-248">ReadItem</span></span>|
|[<span data-ttu-id="1fbec-249">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1fbec-249">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1fbec-250">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1fbec-250">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1fbec-251">例</span><span class="sxs-lookup"><span data-stu-id="1fbec-251">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

<br>

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="1fbec-252">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="1fbec-252">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="1fbec-253">新しい予定を作成するためのフォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="1fbec-253">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="1fbec-254">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="1fbec-254">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="1fbec-p107">`displayNewAppointmentForm` メソッドを使用すると、ユーザーが新しい予定または会議を作成できるフォームが開きます。パラメーターを指定すると、予定のフォーム フィールドにパラメーターの内容が自動的に設定されます。</span><span class="sxs-lookup"><span data-stu-id="1fbec-p107">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="1fbec-p108">Outlook on the web およびモバイル デバイスでは、このメソッドは常に出席者フィールドが含まれるフォームを表示します。入力引数として出席者を指定しないと、このメソッドは **[保存]** ボタンのあるフォームを表示します。出席者を指定した場合には、フォームにその出席者と **[送信]** ボタンが表示されます。</span><span class="sxs-lookup"><span data-stu-id="1fbec-p108">In Outlook on the web and mobile devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="1fbec-p109">Outlook リッチ クライアントと Outlook RT で、`requiredAttendees`、`optionalAttendees`、または `resources` パラメーターに出席者またはリソースを指定し、このメソッドを実行すると、**[送信]** ボタンがある会議フォームが表示されます。受信者を指定せずにこのメソッドを実行すると、**[保存して閉じる]** ボタンがある予定フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="1fbec-p109">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="1fbec-262">パラメーターのいずれかが指定のサイズ制限を超える場合、または不明なパラメーター名が指定されている場合は、例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="1fbec-262">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1fbec-263">パラメーター</span><span class="sxs-lookup"><span data-stu-id="1fbec-263">Parameters</span></span>

|<span data-ttu-id="1fbec-264">名前</span><span class="sxs-lookup"><span data-stu-id="1fbec-264">Name</span></span>| <span data-ttu-id="1fbec-265">種類</span><span class="sxs-lookup"><span data-stu-id="1fbec-265">Type</span></span>| <span data-ttu-id="1fbec-266">説明</span><span class="sxs-lookup"><span data-stu-id="1fbec-266">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="1fbec-267">Object</span><span class="sxs-lookup"><span data-stu-id="1fbec-267">Object</span></span> | <span data-ttu-id="1fbec-268">新しい予定を記述するパラメーターのディクショナリ。</span><span class="sxs-lookup"><span data-stu-id="1fbec-268">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="1fbec-269">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)&gt;</span><span class="sxs-lookup"><span data-stu-id="1fbec-269">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)&gt;</span></span> | <span data-ttu-id="1fbec-p110">予定に必要な各出席者について、メール アドレスを含む文字列の配列、または `EmailAddressDetails` オブジェクトを含む配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="1fbec-p110">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="1fbec-272">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)&gt;</span><span class="sxs-lookup"><span data-stu-id="1fbec-272">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)&gt;</span></span> | <span data-ttu-id="1fbec-p111">予定の各任意出席者について、メール アドレスを含む文字列の配列、または `EmailAddressDetails` オブジェクトを含む配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="1fbec-p111">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="1fbec-275">日付</span><span class="sxs-lookup"><span data-stu-id="1fbec-275">Date</span></span> | <span data-ttu-id="1fbec-276">予定の開始日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="1fbec-276">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="1fbec-277">日付</span><span class="sxs-lookup"><span data-stu-id="1fbec-277">Date</span></span> | <span data-ttu-id="1fbec-278">予定の終了日時を指定する `Date` オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="1fbec-278">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="1fbec-279">String</span><span class="sxs-lookup"><span data-stu-id="1fbec-279">String</span></span> | <span data-ttu-id="1fbec-p112">予定の場所を含む文字列。文字列は最大 255 文字に制限されます。</span><span class="sxs-lookup"><span data-stu-id="1fbec-p112">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="1fbec-282">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="1fbec-282">Array.&lt;String&gt;</span></span> | <span data-ttu-id="1fbec-p113">予定に必要なリソースを含む文字列の配列。配列の上限は 100 エントリです。</span><span class="sxs-lookup"><span data-stu-id="1fbec-p113">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="1fbec-285">String</span><span class="sxs-lookup"><span data-stu-id="1fbec-285">String</span></span> | <span data-ttu-id="1fbec-p114">予定の件名を含む文字列です。文字列は最大 255 文字に制限されます。</span><span class="sxs-lookup"><span data-stu-id="1fbec-p114">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="1fbec-288">String</span><span class="sxs-lookup"><span data-stu-id="1fbec-288">String</span></span> | <span data-ttu-id="1fbec-p115">予定の本文。本文の内容は、最大サイズが 32 KB に制限されます。</span><span class="sxs-lookup"><span data-stu-id="1fbec-p115">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="1fbec-291">要件</span><span class="sxs-lookup"><span data-stu-id="1fbec-291">Requirements</span></span>

|<span data-ttu-id="1fbec-292">要件</span><span class="sxs-lookup"><span data-stu-id="1fbec-292">Requirement</span></span>| <span data-ttu-id="1fbec-293">値</span><span class="sxs-lookup"><span data-stu-id="1fbec-293">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fbec-294">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1fbec-294">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fbec-295">1.0</span><span class="sxs-lookup"><span data-stu-id="1fbec-295">1.0</span></span>|
|[<span data-ttu-id="1fbec-296">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1fbec-296">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fbec-297">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1fbec-297">ReadItem</span></span>|
|[<span data-ttu-id="1fbec-298">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1fbec-298">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1fbec-299">読み取り</span><span class="sxs-lookup"><span data-stu-id="1fbec-299">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1fbec-300">例</span><span class="sxs-lookup"><span data-stu-id="1fbec-300">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="1fbec-301">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="1fbec-301">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="1fbec-302">Exchange Server から添付ファイルやアイテムを取得するために使うトークンを含む文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="1fbec-302">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="1fbec-p116">`getCallbackTokenAsync` メソッドは、ユーザーのメールボックスをホストする Exchange Server から不透明なトークンを取得する非同期の呼び出しを行います。コールバック トークンの有効期間は 5 分です。</span><span class="sxs-lookup"><span data-stu-id="1fbec-p116">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="1fbec-305">トークンと、添付ファイル識別子またはアイテム識別子の両方をサードパーティのシステムに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="1fbec-305">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="1fbec-306">サードパーティシステムは、トークンをベアラー認証トークンとして使用して、Exchange Web サービス (EWS) の[Getattachment](/exchange/client-developer/web-service-reference/getattachment-operation)操作または[GetItem](/exchange/client-developer/web-service-reference/getitem-operation)操作を呼び出して、添付ファイルまたはアイテムを返します。</span><span class="sxs-lookup"><span data-stu-id="1fbec-306">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item.</span></span> <span data-ttu-id="1fbec-307">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="1fbec-307">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="1fbec-308">メソッドを`getCallbackTokenAsync`呼び出すには、 **ReadItem**の最低限のアクセス許可レベルが必要です。</span><span class="sxs-lookup"><span data-stu-id="1fbec-308">Calling the `getCallbackTokenAsync` method requires a minimum permission level of **ReadItem**.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1fbec-309">パラメーター</span><span class="sxs-lookup"><span data-stu-id="1fbec-309">Parameters</span></span>

|<span data-ttu-id="1fbec-310">名前</span><span class="sxs-lookup"><span data-stu-id="1fbec-310">Name</span></span>| <span data-ttu-id="1fbec-311">種類</span><span class="sxs-lookup"><span data-stu-id="1fbec-311">Type</span></span>| <span data-ttu-id="1fbec-312">属性</span><span class="sxs-lookup"><span data-stu-id="1fbec-312">Attributes</span></span>| <span data-ttu-id="1fbec-313">説明</span><span class="sxs-lookup"><span data-stu-id="1fbec-313">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="1fbec-314">function</span><span class="sxs-lookup"><span data-stu-id="1fbec-314">function</span></span>||<span data-ttu-id="1fbec-315">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="1fbec-315">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="1fbec-316">トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="1fbec-316">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="1fbec-317">エラーが発生した場合、 `asyncResult.error` および `asyncResult.diagnostics` のプロパティで追加情報が提供される場合があります。</span><span class="sxs-lookup"><span data-stu-id="1fbec-317">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="1fbec-318">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="1fbec-318">Object</span></span>| <span data-ttu-id="1fbec-319">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="1fbec-319">&lt;optional&gt;</span></span>|<span data-ttu-id="1fbec-320">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="1fbec-320">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="1fbec-321">エラー</span><span class="sxs-lookup"><span data-stu-id="1fbec-321">Errors</span></span>

|<span data-ttu-id="1fbec-322">エラー コード</span><span class="sxs-lookup"><span data-stu-id="1fbec-322">Error code</span></span>|<span data-ttu-id="1fbec-323">説明</span><span class="sxs-lookup"><span data-stu-id="1fbec-323">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="1fbec-324">要求が失敗しました。</span><span class="sxs-lookup"><span data-stu-id="1fbec-324">The request has failed.</span></span> <span data-ttu-id="1fbec-325">HTTP エラーコードの diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="1fbec-325">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="1fbec-326">Exchange サーバーがエラーを返しました。</span><span class="sxs-lookup"><span data-stu-id="1fbec-326">The Exchange server returned an error.</span></span> <span data-ttu-id="1fbec-327">詳細については、diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="1fbec-327">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="1fbec-328">ユーザーはネットワークに接続されていません。</span><span class="sxs-lookup"><span data-stu-id="1fbec-328">The user is no longer connected to the network.</span></span> <span data-ttu-id="1fbec-329">ネットワーク接続を確認し、やり直してください。</span><span class="sxs-lookup"><span data-stu-id="1fbec-329">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1fbec-330">要件</span><span class="sxs-lookup"><span data-stu-id="1fbec-330">Requirements</span></span>

|<span data-ttu-id="1fbec-331">要件</span><span class="sxs-lookup"><span data-stu-id="1fbec-331">Requirement</span></span>| <span data-ttu-id="1fbec-332">値</span><span class="sxs-lookup"><span data-stu-id="1fbec-332">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fbec-333">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1fbec-333">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fbec-334">1.0</span><span class="sxs-lookup"><span data-stu-id="1fbec-334">1.0</span></span>|
|[<span data-ttu-id="1fbec-335">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1fbec-335">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fbec-336">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1fbec-336">ReadItem</span></span>|
|[<span data-ttu-id="1fbec-337">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1fbec-337">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1fbec-338">読み取り</span><span class="sxs-lookup"><span data-stu-id="1fbec-338">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1fbec-339">例</span><span class="sxs-lookup"><span data-stu-id="1fbec-339">Example</span></span>

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

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="1fbec-340">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="1fbec-340">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="1fbec-341">ユーザーと Office アドインを識別するトークンを取得します。</span><span class="sxs-lookup"><span data-stu-id="1fbec-341">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="1fbec-342">`getUserIdentityTokenAsync` メソッドは、[アドインとユーザーをサード パーティのシステムで識別して認証](/outlook/add-ins/authentication)することのできるトークンを返します。</span><span class="sxs-lookup"><span data-stu-id="1fbec-342">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="1fbec-343">パラメーター</span><span class="sxs-lookup"><span data-stu-id="1fbec-343">Parameters</span></span>

|<span data-ttu-id="1fbec-344">名前</span><span class="sxs-lookup"><span data-stu-id="1fbec-344">Name</span></span>| <span data-ttu-id="1fbec-345">種類</span><span class="sxs-lookup"><span data-stu-id="1fbec-345">Type</span></span>| <span data-ttu-id="1fbec-346">属性</span><span class="sxs-lookup"><span data-stu-id="1fbec-346">Attributes</span></span>| <span data-ttu-id="1fbec-347">説明</span><span class="sxs-lookup"><span data-stu-id="1fbec-347">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="1fbec-348">function</span><span class="sxs-lookup"><span data-stu-id="1fbec-348">function</span></span>||<span data-ttu-id="1fbec-349">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="1fbec-349">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="1fbec-350">トークンは、`asyncResult.value` プロパティで文字列として提供されます。</span><span class="sxs-lookup"><span data-stu-id="1fbec-350">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="1fbec-351">エラーが発生した場合、 `asyncResult.error` および `asyncResult.diagnostics` のプロパティで追加情報が提供される場合があります。</span><span class="sxs-lookup"><span data-stu-id="1fbec-351">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="1fbec-352">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="1fbec-352">Object</span></span>| <span data-ttu-id="1fbec-353">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="1fbec-353">&lt;optional&gt;</span></span>|<span data-ttu-id="1fbec-354">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="1fbec-354">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="1fbec-355">エラー</span><span class="sxs-lookup"><span data-stu-id="1fbec-355">Errors</span></span>

|<span data-ttu-id="1fbec-356">エラー コード</span><span class="sxs-lookup"><span data-stu-id="1fbec-356">Error code</span></span>|<span data-ttu-id="1fbec-357">説明</span><span class="sxs-lookup"><span data-stu-id="1fbec-357">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="1fbec-358">要求が失敗しました。</span><span class="sxs-lookup"><span data-stu-id="1fbec-358">The request has failed.</span></span> <span data-ttu-id="1fbec-359">HTTP エラーコードの diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="1fbec-359">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="1fbec-360">Exchange サーバーがエラーを返しました。</span><span class="sxs-lookup"><span data-stu-id="1fbec-360">The Exchange server returned an error.</span></span> <span data-ttu-id="1fbec-361">詳細については、diagnostics オブジェクトを参照してください。</span><span class="sxs-lookup"><span data-stu-id="1fbec-361">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="1fbec-362">ユーザーはネットワークに接続されていません。</span><span class="sxs-lookup"><span data-stu-id="1fbec-362">The user is no longer connected to the network.</span></span> <span data-ttu-id="1fbec-363">ネットワーク接続を確認し、やり直してください。</span><span class="sxs-lookup"><span data-stu-id="1fbec-363">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1fbec-364">要件</span><span class="sxs-lookup"><span data-stu-id="1fbec-364">Requirements</span></span>

|<span data-ttu-id="1fbec-365">要件</span><span class="sxs-lookup"><span data-stu-id="1fbec-365">Requirement</span></span>| <span data-ttu-id="1fbec-366">値</span><span class="sxs-lookup"><span data-stu-id="1fbec-366">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fbec-367">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1fbec-367">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fbec-368">1.0</span><span class="sxs-lookup"><span data-stu-id="1fbec-368">1.0</span></span>|
|[<span data-ttu-id="1fbec-369">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1fbec-369">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fbec-370">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1fbec-370">ReadItem</span></span>|
|[<span data-ttu-id="1fbec-371">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1fbec-371">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1fbec-372">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1fbec-372">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1fbec-373">例</span><span class="sxs-lookup"><span data-stu-id="1fbec-373">Example</span></span>

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

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="1fbec-374">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="1fbec-374">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="1fbec-375">ユーザーのメールボックスをホストしている Exchange サーバー上の Exchange Web サービス (EWS) のサービスに対して非同期の要求を行います。</span><span class="sxs-lookup"><span data-stu-id="1fbec-375">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="1fbec-376">このメソッドは、次のシナリオではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="1fbec-376">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="1fbec-377">Outlook on iOS または Android の場合</span><span class="sxs-lookup"><span data-stu-id="1fbec-377">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="1fbec-378">アドインが Gmail のメールボックスに読み込まれる場合</span><span class="sxs-lookup"><span data-stu-id="1fbec-378">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="1fbec-379">このような場合は、アドインでは [REST API を使用](/outlook/add-ins/use-rest-api)して、代わりにユーザーのメールボックスにアクセスする必要があります。</span><span class="sxs-lookup"><span data-stu-id="1fbec-379">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="1fbec-p124">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange. See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span><span class="sxs-lookup"><span data-stu-id="1fbec-p124">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange. See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="1fbec-382">`makeEwsRequestAsync` メソッドでは、フォルダー関連アイテムを要求できません。</span><span class="sxs-lookup"><span data-stu-id="1fbec-382">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="1fbec-383">XML 要求では UTF-8 エンコードを指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="1fbec-383">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="1fbec-p125">`makeEwsRequestAsync` メソッドを使用するには、アドインに **ReadWriteMailbox** アクセス許可が必要です。**ReadWriteMailbox** アクセス許可と、`makeEwsRequestAsync` メソッドで呼び出せる EWS 操作の使い方については、「[ユーザーのメールボックスへのメール アドイン アクセスのアクセス許可を指定する](/outlook/add-ins/understanding-outlook-add-in-permissions)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1fbec-p125">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="1fbec-386">サーバー管理者は、クライアント アクセス サーバーの EWS ディレクトリで `OAuthAuthentication` を true に設定して、`makeEwsRequestAsync` メソッドで EWS 要求を行うことができるようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="1fbec-386">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="1fbec-387">バージョンの相違点</span><span class="sxs-lookup"><span data-stu-id="1fbec-387">Version differences</span></span>

<span data-ttu-id="1fbec-388">バージョン 15.0.4535.1004 より前のバージョンの Outlook で実行しているメール アプリで `makeEwsRequestAsync` メソッドを使う場合は、エンコード値を `ISO-8859-1` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="1fbec-388">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="1fbec-p126">Outlook on the web でメール アプリを実行している場合は、エンコード値を設定する必要はありません。mailbox.diagnostics.hostName プロパティを使って、メール アプリを Outlook で実行しているのか、Outlook on the web で実行しているのかを確認できます。mailbox.diagnostics.hostVersion プロパティを使って、どのバージョンの Outlook を使って実行しているかを確認できます。</span><span class="sxs-lookup"><span data-stu-id="1fbec-p126">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1fbec-392">パラメーター</span><span class="sxs-lookup"><span data-stu-id="1fbec-392">Parameters</span></span>

|<span data-ttu-id="1fbec-393">名前</span><span class="sxs-lookup"><span data-stu-id="1fbec-393">Name</span></span>| <span data-ttu-id="1fbec-394">種類</span><span class="sxs-lookup"><span data-stu-id="1fbec-394">Type</span></span>| <span data-ttu-id="1fbec-395">属性</span><span class="sxs-lookup"><span data-stu-id="1fbec-395">Attributes</span></span>| <span data-ttu-id="1fbec-396">説明</span><span class="sxs-lookup"><span data-stu-id="1fbec-396">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="1fbec-397">String</span><span class="sxs-lookup"><span data-stu-id="1fbec-397">String</span></span>||<span data-ttu-id="1fbec-398">EWS 要求です。</span><span class="sxs-lookup"><span data-stu-id="1fbec-398">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="1fbec-399">function</span><span class="sxs-lookup"><span data-stu-id="1fbec-399">function</span></span>||<span data-ttu-id="1fbec-400">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="1fbec-400">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="1fbec-p127">The XML result of the EWS call is provided as a string in the `asyncResult.value` property. If the result exceeds 1 MB in size, an error message is returned instead.</span><span class="sxs-lookup"><span data-stu-id="1fbec-p127">The XML result of the EWS call is provided as a string in the `asyncResult.value` property. If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="1fbec-403">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="1fbec-403">Object</span></span>| <span data-ttu-id="1fbec-404">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="1fbec-404">&lt;optional&gt;</span></span>|<span data-ttu-id="1fbec-405">非同期メソッドに渡される状態データです。</span><span class="sxs-lookup"><span data-stu-id="1fbec-405">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1fbec-406">要件</span><span class="sxs-lookup"><span data-stu-id="1fbec-406">Requirements</span></span>

|<span data-ttu-id="1fbec-407">要件</span><span class="sxs-lookup"><span data-stu-id="1fbec-407">Requirement</span></span>| <span data-ttu-id="1fbec-408">値</span><span class="sxs-lookup"><span data-stu-id="1fbec-408">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fbec-409">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1fbec-409">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fbec-410">1.0</span><span class="sxs-lookup"><span data-stu-id="1fbec-410">1.0</span></span>|
|[<span data-ttu-id="1fbec-411">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1fbec-411">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fbec-412">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="1fbec-412">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="1fbec-413">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1fbec-413">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1fbec-414">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1fbec-414">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1fbec-415">例</span><span class="sxs-lookup"><span data-stu-id="1fbec-415">Example</span></span>

<span data-ttu-id="1fbec-416">次の例は、`makeEwsRequestAsync` を呼び出し、`GetItem` 操作を使って項目の件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="1fbec-416">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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
