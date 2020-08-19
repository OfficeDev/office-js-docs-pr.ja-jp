---
title: Outlook アドインで定期的なアイテムを取得して設定する
description: このトピックでは、Office JavaScript API を使用して、Outlook のアドインでさまざまな定期的なアイテムのプロパティを取得および設定する方法を示します。
ms.date: 08/18/2020
localization_priority: Normal
ms.openlocfilehash: 0b179725677f071fe2ae7baf1c719add5ccd8aa7
ms.sourcegitcommit: e9f23a2857b90a7c17e3152292b548a13a90aa33
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/19/2020
ms.locfileid: "46803745"
---
# <a name="get-and-set-recurrence"></a><span data-ttu-id="eecee-103">定期的なアイテムを取得および設定する</span><span class="sxs-lookup"><span data-stu-id="eecee-103">Get and set recurrence</span></span>

<span data-ttu-id="eecee-104">毎週のチーム プロジェクトの進捗会議や毎年の誕生日通知など、定期的な予定の作成や更新が必要な場合があります。</span><span class="sxs-lookup"><span data-stu-id="eecee-104">Sometimes you need to create and update a recurring appointment, such as a weekly status meeting for a team project or a yearly birthday reminder.</span></span> <span data-ttu-id="eecee-105">Office JavaScript API を使用して、アドイン内の定期的な予定のパターンを管理することができます。</span><span class="sxs-lookup"><span data-stu-id="eecee-105">You can use the Office JavaScript API to manage the recurrence patterns of an appointment series in your add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="eecee-106">この機能のサポートは、要件セット1.7 で導入されました。</span><span class="sxs-lookup"><span data-stu-id="eecee-106">Support for this feature was introduced in requirement set 1.7.</span></span> <span data-ttu-id="eecee-107">この要件セットをサポートする [クライアントおよびプラットフォーム](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) を参照してください。</span><span class="sxs-lookup"><span data-stu-id="eecee-107">See [clients and platforms](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

## <a name="available-recurrence-patterns"></a><span data-ttu-id="eecee-108">使用可能な定期的なパターン</span><span class="sxs-lookup"><span data-stu-id="eecee-108">Available recurrence patterns</span></span>

<span data-ttu-id="eecee-109">定期的なパターンを構成するには、[定期的なパターン](/javascript/api/outlook/office.mailboxenums.recurrencetype)と、該当する[定期的なアイテムのプロパティ](/javascript/api/outlook/office.recurrenceproperties) (ある場合) を結合する必要があります。</span><span class="sxs-lookup"><span data-stu-id="eecee-109">To configure the recurrence pattern, you need to combine the [recurrence type](/javascript/api/outlook/office.mailboxenums.recurrencetype) and its applicable [recurrence properties](/javascript/api/outlook/office.recurrenceproperties) (if any).</span></span>

<span data-ttu-id="eecee-110">**表 1. 定期的なパターンと、適用可能なプロパティ**</span><span class="sxs-lookup"><span data-stu-id="eecee-110">**Table 1. Recurrence types and their applicable properties**</span></span>

|<span data-ttu-id="eecee-111">定期的なパターン</span><span class="sxs-lookup"><span data-stu-id="eecee-111">Recurrence type</span></span>|<span data-ttu-id="eecee-112">有効な定期的なアイテムのプロパティ</span><span class="sxs-lookup"><span data-stu-id="eecee-112">Valid recurrence properties</span></span>|<span data-ttu-id="eecee-113">使用方法</span><span class="sxs-lookup"><span data-stu-id="eecee-113">Usage</span></span>|
|---|---|---|
|`daily`|-&nbsp;[`interval`][interval link]|<span data-ttu-id="eecee-114">*interval* 日に一度、予定が発生する。</span><span class="sxs-lookup"><span data-stu-id="eecee-114">An appointment occurs every *interval* days.</span></span> <span data-ttu-id="eecee-115">例: 予定が **_2 日_** おきに発生する。</span><span class="sxs-lookup"><span data-stu-id="eecee-115">Example: An appointment occurs every **_2_** days.</span></span>|
|`weekday`|<span data-ttu-id="eecee-116">なし。</span><span class="sxs-lookup"><span data-stu-id="eecee-116">None.</span></span>|<span data-ttu-id="eecee-117">予定が平日に毎日発生する。</span><span class="sxs-lookup"><span data-stu-id="eecee-117">An appointment occurs every weekday.</span></span>|
|`monthly`|-&nbsp;[`interval`][interval link]<br/>-&nbsp;[`dayOfMonth`][dayOfMonth link]<br/>-&nbsp;[`dayOfWeek`][dayOfWeek link]<br/>-&nbsp;[`weekNumber`][weekNumber link]|<span data-ttu-id="eecee-118">- 予定が *interval* か月に一度、*dayOfMonth* 日に発生する。</span><span class="sxs-lookup"><span data-stu-id="eecee-118">- An appointment occurs on day *dayOfMonth* every *interval* months.</span></span> <span data-ttu-id="eecee-119">例: 予定が \*\*_4_ **か月に一度、**_5_ \*\*日に発生する。</span><span class="sxs-lookup"><span data-stu-id="eecee-119">Example: An appointment occurs on day **_5_** every **_4_** months.</span></span><br/><br/><span data-ttu-id="eecee-120">- 予定が、*interval* か月に一度、第 *weekNumber* 週の *dayOfWeek* 日に発生する。</span><span class="sxs-lookup"><span data-stu-id="eecee-120">- An appointment occurs on the *weekNumber* *dayOfWeek* every *interval* months.</span></span> <span data-ttu-id="eecee-121">例: 予定が、**_2_** か月に一度、第 **_3_** **_木曜日_** に発生する。</span><span class="sxs-lookup"><span data-stu-id="eecee-121">Example: An appointment occurs on the **_third_** **_Thursday_** every **_2_** months.</span></span>|
|`weekly`|-&nbsp;[`interval`][interval link]<br/>-&nbsp;[`days`][days link]|<span data-ttu-id="eecee-122">予定が *interval* 週間に一度、*days* に発生する。</span><span class="sxs-lookup"><span data-stu-id="eecee-122">An appointment occurs on *days* every *interval* weeks.</span></span> <span data-ttu-id="eecee-123">例: 予定が\*\* _2_ **週間に一度、**_火曜日_と_木曜日_\*\* に発生する。</span><span class="sxs-lookup"><span data-stu-id="eecee-123">Example: An appointment occurs on **_Tuesday_ and _Thursday_** every **_2_** weeks.</span></span>|
|`yearly`|-&nbsp;[`interval`][interval link]<br/>-&nbsp;[`dayOfMonth`][dayOfMonth link]<br/>-&nbsp;[`dayOfWeek`][dayOfWeek link]<br/>-&nbsp;[`weekNumber`][weekNumber link]<br/>-&nbsp;[`month`][month link]|<span data-ttu-id="eecee-124">- 予定が、*interval* 年に一度、*month* の *dayOfMonth* 日に発生する。</span><span class="sxs-lookup"><span data-stu-id="eecee-124">- An appointment occurs on day *dayOfMonth* of *month* every *interval* years.</span></span> <span data-ttu-id="eecee-125">例: 予定が **_4_** 年に一度、**_9 月_** **_7_** 日に発生する。</span><span class="sxs-lookup"><span data-stu-id="eecee-125">Example: An appointment occurs on day **_7_** of **_September_** every **_4_** years.</span></span><br/><br/><span data-ttu-id="eecee-126">- 予定が、*interval* 年に一度、*month* の第 *weekNumber* 週の *dayOfWeek* に発生する。</span><span class="sxs-lookup"><span data-stu-id="eecee-126">- An appointment occurs on the *weekNumber* *dayOfWeek* of *month* every *interval* years.</span></span> <span data-ttu-id="eecee-127">例: 予定が、**_2_** 年に一度、**_9 月_** の**_最初_** の**_木曜日_** に発生する。</span><span class="sxs-lookup"><span data-stu-id="eecee-127">Example: An appointment occurs on the **_first_** **_Thursday_** of **_September_** every **_2_** years.</span></span>|

> [!NOTE]
> <span data-ttu-id="eecee-128"> の定期的なパターンで [`firstDayOfWeek`][firstDayOfWeek link]`weekly` プロパティを使用することもできます。</span><span class="sxs-lookup"><span data-stu-id="eecee-128">You can also use the [`firstDayOfWeek`][firstDayOfWeek link] property with the `weekly` recurrence type.</span></span> <span data-ttu-id="eecee-129">指定された日は定期的なアイテムのダイアログに表示された日にちのリストを開始させます。</span><span class="sxs-lookup"><span data-stu-id="eecee-129">The specified day will start the list of days displayed in the recurrence dialog.</span></span>

## <a name="access-recurrence"></a><span data-ttu-id="eecee-130">定期的なアイテムにアクセスする</span><span class="sxs-lookup"><span data-stu-id="eecee-130">Access recurrence</span></span>

<span data-ttu-id="eecee-131">予定の開催者であるか出席者であるかによって、定期的なパターンにアクセスする方法、およびアクセスしてできることが変わります。</span><span class="sxs-lookup"><span data-stu-id="eecee-131">How you access the recurrence pattern and what you can do with it depends on if you're the appointment organizer or an attendee.</span></span>

<span data-ttu-id="eecee-132">**表 2. 適用可能な予定の状態**</span><span class="sxs-lookup"><span data-stu-id="eecee-132">**Table 2. Applicable appointment states**</span></span>

|<span data-ttu-id="eecee-133">予定の状態</span><span class="sxs-lookup"><span data-stu-id="eecee-133">Appointment state</span></span>|<span data-ttu-id="eecee-134">編集可能な定期的なアイテムですか。</span><span class="sxs-lookup"><span data-stu-id="eecee-134">Is recurrence editable?</span></span>|<span data-ttu-id="eecee-135">表示可能な定期的なアイテムですか。</span><span class="sxs-lookup"><span data-stu-id="eecee-135">Is recurrence viewable?</span></span>|
|---|---|---|
|<span data-ttu-id="eecee-136">予定の開催者 - 定期的な予定を作成する</span><span class="sxs-lookup"><span data-stu-id="eecee-136">Appointment organizer - compose series</span></span>|<span data-ttu-id="eecee-137">はい ([`setAsync`][setAsync link])</span><span class="sxs-lookup"><span data-stu-id="eecee-137">Yes ([`setAsync`][setAsync link])</span></span>|<span data-ttu-id="eecee-138">はい ([`getAsync`][getAsync link])</span><span class="sxs-lookup"><span data-stu-id="eecee-138">Yes ([`getAsync`][getAsync link])</span></span>|
|<span data-ttu-id="eecee-139">予定の開催者 - インスタンスを作成する</span><span class="sxs-lookup"><span data-stu-id="eecee-139">Appointment organizer - compose instance</span></span>|<span data-ttu-id="eecee-140">いいえ (`setAsync` がエラーを返します)</span><span class="sxs-lookup"><span data-stu-id="eecee-140">No (`setAsync` returns an error)</span></span>|<span data-ttu-id="eecee-141">はい ([`getAsync`][getAsync link])</span><span class="sxs-lookup"><span data-stu-id="eecee-141">Yes ([`getAsync`][getAsync link])</span></span>|
|<span data-ttu-id="eecee-142">予定の出席者 - 定期的な予定を確認する</span><span class="sxs-lookup"><span data-stu-id="eecee-142">Appointment attendee - read series</span></span>|<span data-ttu-id="eecee-143">いいえ (`setAsync` が使用不可)</span><span class="sxs-lookup"><span data-stu-id="eecee-143">No (`setAsync` not available)</span></span>|<span data-ttu-id="eecee-144">はい ([`item.recurrence`][item.recurrence link])</span><span class="sxs-lookup"><span data-stu-id="eecee-144">Yes ([`item.recurrence`][item.recurrence link])</span></span>|
|<span data-ttu-id="eecee-145">予定の出席者 - インスタンスを読む</span><span class="sxs-lookup"><span data-stu-id="eecee-145">Appointment attendee - read instance</span></span>|<span data-ttu-id="eecee-146">いいえ (`setAsync` が使用不可)</span><span class="sxs-lookup"><span data-stu-id="eecee-146">No (`setAsync` not available)</span></span>|<span data-ttu-id="eecee-147">はい ([`item.recurrence`][item.recurrence link])</span><span class="sxs-lookup"><span data-stu-id="eecee-147">Yes ([`item.recurrence`][item.recurrence link])</span></span>|
|<span data-ttu-id="eecee-148">会議出席依頼 - 定期的な予定を確認する</span><span class="sxs-lookup"><span data-stu-id="eecee-148">Meeting request - read series</span></span>|<span data-ttu-id="eecee-149">いいえ (`setAsync` が使用不可)</span><span class="sxs-lookup"><span data-stu-id="eecee-149">No (`setAsync` not available)</span></span>|<span data-ttu-id="eecee-150">はい ([`item.recurrence`][item.recurrence link])</span><span class="sxs-lookup"><span data-stu-id="eecee-150">Yes ([`item.recurrence`][item.recurrence link])</span></span>|
|<span data-ttu-id="eecee-151">会議出席依頼 - インスタンスを確認する</span><span class="sxs-lookup"><span data-stu-id="eecee-151">Meeting request - read instance</span></span>|<span data-ttu-id="eecee-152">いいえ (`setAsync` が使用不可)</span><span class="sxs-lookup"><span data-stu-id="eecee-152">No (`setAsync` not available)</span></span>|<span data-ttu-id="eecee-153">はい ([`item.recurrence`][item.recurrence link])</span><span class="sxs-lookup"><span data-stu-id="eecee-153">Yes ([`item.recurrence`][item.recurrence link])</span></span>|

## <a name="set-recurrence-as-the-organizer"></a><span data-ttu-id="eecee-154">定期的なアイテムを開催者として設定する</span><span class="sxs-lookup"><span data-stu-id="eecee-154">Set recurrence as the organizer</span></span>

<span data-ttu-id="eecee-155">定期的なパターンを使用するには、定期的な予定の開始日時、終了日時も決定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="eecee-155">Along with the recurrence pattern, you also need to determine the start and end dates and times of your appointment series.</span></span> <span data-ttu-id="eecee-156">[`SeriesTime`][SeriesTime link] オブジェクトはその情報を管理するために使用します。</span><span class="sxs-lookup"><span data-stu-id="eecee-156">The [`SeriesTime`][SeriesTime link] object is used to manage that information.</span></span>

<span data-ttu-id="eecee-157">予定の開催者は、作成モードでのみ、定期的な予定のパターンを指定できます。</span><span class="sxs-lookup"><span data-stu-id="eecee-157">The appointment organizer can specify the recurrence pattern for an appointment series in compose mode only.</span></span> <span data-ttu-id="eecee-158">次の例では、2019 年 11 月 2 日から 2019 年 12 月 2 日の期間中に、毎週火曜日と木曜日の、午前 10 時 30 分から午前 11 時 00 分 (米国太平洋標準時) に発生する定期的な予定が設定されています。</span><span class="sxs-lookup"><span data-stu-id="eecee-158">In the following example, the appointment series is set to occur from 10:30 AM to 11:00 AM PST every Tuesday and Thursday during the period November 2, 2019 to December 2, 2019.</span></span>

```js
var seriesTimeObject = new Office.SeriesTime();
seriesTimeObject.setStartDate(2019,10,2);
seriesTimeObject.setEndDate(2019,11,2);
seriesTimeObject.setStartTime(10,30);
seriesTimeObject.setDuration(30);

var pattern = {
    "seriesTime": seriesTimeObject,
    "recurrenceType": "weekly",
    "recurrenceProperties": {"interval": 1, "days": ["tue", "thu"]},
    "recurrenceTimeZone": {"name": "Pacific Standard Time"}};

Office.context.mailbox.item.recurrence.setAsync(pattern, callback);

function callback(asyncResult)
{
    console.log(JSON.stringify(asyncResult));
}
```

## <a name="change-recurrence-as-the-organizer"></a><span data-ttu-id="eecee-159">開催者として定期的なアイテムを変更する</span><span class="sxs-lookup"><span data-stu-id="eecee-159">Change recurrence as the organizer</span></span>

<span data-ttu-id="eecee-160">次の例では、作成モードでは、予定の開催者が一連の定期的な予定またはその系列のインスタンスを指定して、定期的な予定のオブジェクトを取得し、新しい定期的なアイテムの期間を設定します。</span><span class="sxs-lookup"><span data-stu-id="eecee-160">In the following example, in compose mode, the appointment organizer gets the recurrence object of an appointment series given the series or an instance of that series, then sets a new recurrence duration.</span></span>

```js
Office.context.mailbox.item.recurrence.getAsync(callback);

function callback(asyncResult) {
  var recurrencePattern = asyncResult.value;
  recurrencePattern.seriesTime.setDuration(60);
  Office.context.mailbox.item.recurrence.setAsync(recurrencePattern, (asyncResult) => {
    if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
      console.log("failed");
      return;
    }

    console.log("success");
  });
}
```

## <a name="get-recurrence"></a><span data-ttu-id="eecee-161">定期的なアイテムを取得する</span><span class="sxs-lookup"><span data-stu-id="eecee-161">Get recurrence</span></span>

### <a name="get-recurrence-as-the-organizer"></a><span data-ttu-id="eecee-162">定期的なアイテムを開催者として取得する</span><span class="sxs-lookup"><span data-stu-id="eecee-162">Get recurrence as the organizer</span></span>

<span data-ttu-id="eecee-163">次の例では、予定の開催者が作成モードで、定期的な予定やそのインスタンスで、その予定の繰り返しオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="eecee-163">In the following example, in compose mode, the appointment organizer gets the recurrence object of an appointment series given the series or an instance of that series.</span></span>

```js
Office.context.mailbox.item.recurrence.getAsync(callback);

function callback(asyncResult){
    var context = asyncResult.context;
    var recurrence = asyncResult.value;

    if (recurrence == null) {
        console.log("Non-recurring meeting");
    } else {
        console.log(JSON.stringify(recurrence));
    }
}
```

<span data-ttu-id="eecee-164">次の例では、定期的な予定の繰り返しを取得する `getAsync` コールの結果を表示しています。</span><span class="sxs-lookup"><span data-stu-id="eecee-164">The following example shows the results of the `getAsync` call that retrieves the recurrence for a series.</span></span>

> [!NOTE]
> <span data-ttu-id="eecee-165">この例では、`seriesTimeObject` は `recurrence.seriesTime` プロパティを表す JSON のプレースホルダーです。</span><span class="sxs-lookup"><span data-stu-id="eecee-165">In this example, `seriesTimeObject` is a placeholder for the JSON representing the `recurrence.seriesTime` property.</span></span> <span data-ttu-id="eecee-166">定期的な予定の日時のプロパティを取得するには、[`SeriesTime`][SeriesTime link] メソッドを使用します。</span><span class="sxs-lookup"><span data-stu-id="eecee-166">You should use the [`SeriesTime`][SeriesTime link] methods to get the recurrence date and time properties.</span></span>

```json
{
    "recurrenceType": "weekly",
    "recurrenceProperties": {
        "interval": 1,
        "days": ["tue","thu"],
        "firstDayOfWeek": "sun"},
    "seriesTime": {seriesTimeObject},
    "recurrenceTimeZone": {
        "name": "Pacific Standard Time",
        "offset": -480}}
```

### <a name="get-recurrence-as-an-attendee"></a><span data-ttu-id="eecee-167">定期的なアイテムを出席者として取得する</span><span class="sxs-lookup"><span data-stu-id="eecee-167">Get recurrence as an attendee</span></span>

<span data-ttu-id="eecee-168">次の例では、予定の出席者が、定期的な予定やその予定のインスタンス、または会議出席依頼によって定期的な予定の繰り返しオブジェクトを取得できます。</span><span class="sxs-lookup"><span data-stu-id="eecee-168">In the following example, an appointment attendee can get the recurrence object of an appointment series given the series, an instance of that series, or a meeting request.</span></span>

```js
outputRecurrence(Office.context.mailbox.item);

function outputRecurrence(item) {
    var recurrence = item.recurrence;
    var seriesId = item.seriesId;

    if (recurrence == null) {
        console.log("Non-recurring item");
    } else {
        console.log(JSON.stringify(recurrence));
    }
}
```

<span data-ttu-id="eecee-169">次の例では、定期的な予定の `item.recurrence` プロパティの値を示しています。</span><span class="sxs-lookup"><span data-stu-id="eecee-169">The following example shows the value of the `item.recurrence` property for an appointment series.</span></span>

> [!NOTE]
> <span data-ttu-id="eecee-170">この例では、`seriesTimeObject` は `recurrence.seriesTime` プロパティを表す JSON のプレースホルダーです。</span><span class="sxs-lookup"><span data-stu-id="eecee-170">In this example, `seriesTimeObject` is a placeholder for the JSON representing the `recurrence.seriesTime` property.</span></span> <span data-ttu-id="eecee-171">定期的な予定の日時のプロパティを取得するには、[`SeriesTime`][SeriesTime link] メソッドを使用します。</span><span class="sxs-lookup"><span data-stu-id="eecee-171">You should use the [`SeriesTime`][SeriesTime link] methods to get the recurrence date and time properties.</span></span>

```json
{
    "recurrenceType": "weekly",
    "recurrenceProperties": {
        "interval": 1,
        "days": ["tue","thu"],
        "firstDayOfWeek": "sun"},
    "seriesTime": {seriesTimeObject},
    "recurrenceTimeZone": {
        "name": "Pacific Standard Time",
        "offset": -480}}
```

### <a name="get-the-recurrence-details"></a><span data-ttu-id="eecee-172">定期的なアイテムの詳細を取得する</span><span class="sxs-lookup"><span data-stu-id="eecee-172">Get the recurrence details</span></span>

<span data-ttu-id="eecee-173">(`getAsync` コールバックまたは `item.recurrence` のいずれかから) 繰り返しオブジェクトを取得した後、特定の定期的なアイテムのプロパティを表示できます。</span><span class="sxs-lookup"><span data-stu-id="eecee-173">After you've retrieved the recurrence object (either from the `getAsync` callback or from `item.recurrence`), you can get specific properties of the recurrence.</span></span> <span data-ttu-id="eecee-174">たとえば、 プロパティの[メソッド][SeriesTime link]`recurrence.seriesTime`を使用して定期的なアイテムの開始日時と終了日時を取得できます。</span><span class="sxs-lookup"><span data-stu-id="eecee-174">For example, you can get the start and end dates and times of the series by using [methods][SeriesTime link] on the `recurrence.seriesTime` property.</span></span>

```js
// Get series date and time info
var seriesTime = recurrence.seriesTime;
var startTime = recurrence.seriesTime.getStartTime();
var endTime = recurrence.seriesTime.getEndTime();
var startDate = recurrence.seriesTime.getStartDate();
var endDate = recurrence.seriesTime.getEndDate();
var duration = recurrence.seriesTime.getDuration();

// Get series time zone
var timeZone = recurrence.recurrenceTimeZone;

// Get recurrence properties
var recurrenceProperties = recurrence.recurrenceProperties;

// Get recurrence type
var recurrenceType = recurrence.recurrenceType;
```

## <a name="see-also"></a><span data-ttu-id="eecee-175">関連項目</span><span class="sxs-lookup"><span data-stu-id="eecee-175">See also</span></span>

[<span data-ttu-id="eecee-176">RecurrenceChanged イベント</span><span class="sxs-lookup"><span data-stu-id="eecee-176">RecurrenceChanged event</span></span>](/javascript/api/office/office.eventtype)

[getAsync link]: /javascript/api/outlook/office.recurrence#getasync-options--callback-
[item.recurrence link]: ../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties
[setAsync link]: /javascript/api/outlook/office.recurrence#setasync-recurrencepattern--options--callback-

[dayOfMonth link]: /javascript/api/outlook/office.recurrenceproperties#dayofmonth
[dayOfWeek link]: /javascript/api/outlook/office.recurrenceproperties#dayofweek
[days link]: /javascript/api/outlook/office.recurrenceproperties#days
[firstDayOfWeek link]: /javascript/api/outlook/office.recurrenceproperties#firstdayofweek
[interval link]: /javascript/api/outlook/office.recurrenceproperties#interval
[month link]: /javascript/api/outlook/office.recurrenceproperties#month
[weekNumber link]: /javascript/api/outlook/office.recurrenceproperties#weeknumber

[SeriesTime link]: /javascript/api/outlook/office.seriestime
