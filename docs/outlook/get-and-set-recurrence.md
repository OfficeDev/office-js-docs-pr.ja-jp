---
title: Outlook アドインで定期的なアイテムを取得して設定する
description: このトピックでは、Office JavaScript API を使用して、Outlook のアドインでさまざまな定期的なアイテムのプロパティを取得および設定する方法を示します。
ms.date: 07/08/2022
ms.localizationpriority: medium
ms.openlocfilehash: de6bf8722578729159078396bb2f14a52cb648a3
ms.sourcegitcommit: d8ea4b761f44d3227b7f2c73e52f0d2233bf22e2
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/11/2022
ms.locfileid: "66713113"
---
# <a name="get-and-set-recurrence"></a>定期的なアイテムを取得および設定する

毎週のチーム プロジェクトの進捗会議や毎年の誕生日通知など、定期的な予定の作成や更新が必要な場合があります。 Office JavaScript API を使用して、アドイン内の予定シリーズの繰り返しパターンを管理できます。

> [!NOTE]
> この機能のサポートは [、要件セット 1.7](/javascript/api/requirement-sets/outlook/requirement-set-1.7/outlook-requirement-set-1.7) で導入されました。 この要件セットをサポートする [クライアントおよびプラットフォーム](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients) を参照してください。

## <a name="available-recurrence-patterns"></a>使用可能な定期的なパターン

定期的なパターンを構成するには、[定期的なパターン](/javascript/api/outlook/office.mailboxenums.recurrencetype)と、該当する[定期的なアイテムのプロパティ](/javascript/api/outlook/office.recurrenceproperties) (ある場合) を結合する必要があります。

**表 1. 定期的なパターンと、適用可能なプロパティ**

|定期的なパターン|有効な定期的なアイテムのプロパティ|使用方法|
|---|---|---|
|`daily`|-&nbsp;[`interval`][間隔リンク]|*interval* 日に一度、予定が発生する。 例: 予定が **_2 日_** おきに発生する。|
|`weekday`|なし。|予定が平日に毎日発生する。|
|`monthly`|-&nbsp;[`interval`][間隔リンク]<br/>-&nbsp;[`dayOfMonth`][][dayOfMonth link]<br/>-&nbsp;[`dayOfWeek`][dayOfWeek リンク]<br/>-&nbsp;[`weekNumber`][][weekNumber link]|- 予定が *interval* か月に一度、*dayOfMonth* 日に発生する。 例: 予定が **_4_** か月に一度、**_5_** 日に発生する。<br/><br/>- 予定が、*interval* か月に一度、第 *weekNumber* 週の *dayOfWeek* 日に発生する。 例: 予定が、**_2_** か月に一度、第 **_3_** **_木曜日_** に発生する。|
|`weekly`|-&nbsp;[`interval`][間隔リンク]<br/>-&nbsp;[`days`][日のリンク]|予定が *interval* 週間に一度、*days* に発生する。 例: 予定が **_2_** 週間に一度、**_火曜日_ と _木曜日_** に発生する。|
|`yearly`|-&nbsp;[`interval`][間隔リンク]<br/>-&nbsp;[`dayOfMonth`][][dayOfMonth link]<br/>-&nbsp;[`dayOfWeek`][dayOfWeek リンク]<br/>-&nbsp;[`weekNumber`][][weekNumber link]<br/>-&nbsp;[`month`][月のリンク]|- 予定が、*interval* 年に一度、*month* の *dayOfMonth* 日に発生する。 例: 予定が **_4_** 年に一度、**_9 月_** **_7_** 日に発生する。<br/><br/>- 予定が、*interval* 年に一度、*month* の第 *weekNumber* 週の *dayOfWeek* に発生する。 例: 予定が、**_2_** 年に一度、**_9 月_** の **_最初_** の **_木曜日_** に発生する。|

> [!NOTE]
> [`firstDayOfWeek`][firstDayOfWeek link] プロパティを繰り返しの型と共に `weekly` 使用することもできます。 指定された日は定期的なアイテムのダイアログに表示された日にちのリストを開始させます。

## <a name="access-recurrence"></a>定期的なアイテムにアクセスする

予定の開催者であるか出席者であるかによって、定期的なパターンにアクセスする方法、およびアクセスしてできることが変わります。

**表 2. 適用可能な予定の状態**

|予定の状態|編集可能な定期的なアイテムですか。|表示可能な定期的なアイテムですか。|
|---|---|---|
|予定の開催者 - 定期的な予定を作成する|はい ([`setAsync`][setAsync link])|はい ([`getAsync`][getAsync link])|
|予定の開催者 - インスタンスを作成する|いいえ (`setAsync` がエラーを返します)|はい ([`getAsync`][getAsync link])|
|予定の出席者 - 定期的な予定を確認する|いいえ (`setAsync` が使用不可)|はい ([`item.recurrence`][item.recurrence link])|
|予定の出席者 - インスタンスを読む|いいえ (`setAsync` が使用不可)|はい ([`item.recurrence`][item.recurrence link])|
|会議出席依頼 - 定期的な予定を確認する|いいえ (`setAsync` が使用不可)|はい ([`item.recurrence`][item.recurrence link])|
|会議出席依頼 - インスタンスを確認する|いいえ (`setAsync` が使用不可)|はい ([`item.recurrence`][item.recurrence link])|

## <a name="set-recurrence-as-the-organizer"></a>定期的なアイテムを開催者として設定する

定期的なパターンを使用するには、定期的な予定の開始日時、終了日時も決定する必要があります。 [`SeriesTime`][SeriesTime link] オブジェクトは、その情報を管理するために使用されます。

予定の開催者は、作成モードでのみ、定期的な予定のパターンを指定できます。 次の例では、2019 年 11 月 2 日から 2019 年 12 月 2 日の期間中に、毎週火曜日と木曜日の、午前 10 時 30 分から午前 11 時 00 分 (米国太平洋標準時) に発生する定期的な予定が設定されています。

```js
const seriesTimeObject = new Office.SeriesTime();
seriesTimeObject.setStartDate(2019,10,2);
seriesTimeObject.setEndDate(2019,11,2);
seriesTimeObject.setStartTime(10,30);
seriesTimeObject.setDuration(30);

const pattern = {
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

## <a name="change-recurrence-as-the-organizer"></a>開催者として繰り返しを変更する

次の例では、作成モードでは、予定の開催者は、系列またはその系列のインスタンスを指定して、予定系列の繰り返しオブジェクトを取得し、新しい繰り返し期間を設定します。

```js
Office.context.mailbox.item.recurrence.getAsync(callback);

function callback(asyncResult) {
  const recurrencePattern = asyncResult.value;
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

## <a name="get-recurrence"></a>定期的なアイテムを取得する

### <a name="get-recurrence-as-the-organizer"></a>定期的なアイテムを開催者として取得する

次の例では、予定の開催者が作成モードで、定期的な予定やそのインスタンスで、その予定の繰り返しオブジェクトを取得します。

```js
Office.context.mailbox.item.recurrence.getAsync(callback);

function callback(asyncResult){
    const context = asyncResult.context;
    const recurrence = asyncResult.value;

    if (recurrence == null) {
        console.log("Non-recurring meeting");
    } else {
        console.log(JSON.stringify(recurrence));
    }
}
```

次の例では、定期的な予定の繰り返しを取得する `getAsync` コールの結果を表示しています。

> [!NOTE]
> この例では、`seriesTimeObject` は `recurrence.seriesTime` プロパティを表す JSON のプレースホルダーです。 [`SeriesTime`][SeriesTime link] メソッドを使用して、繰り返しの日付と時刻のプロパティを取得する必要があります。

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

### <a name="get-recurrence-as-an-attendee"></a>定期的なアイテムを出席者として取得する

次の例では、予定の出席者が、定期的な予定やその予定のインスタンス、または会議出席依頼によって定期的な予定の繰り返しオブジェクトを取得できます。

```js
outputRecurrence(Office.context.mailbox.item);

function outputRecurrence(item) {
    const recurrence = item.recurrence;
    const seriesId = item.seriesId;

    if (recurrence == null) {
        console.log("Non-recurring item");
    } else {
        console.log(JSON.stringify(recurrence));
    }
}
```

次の例では、定期的な予定の `item.recurrence` プロパティの値を示しています。

> [!NOTE]
> この例では、`seriesTimeObject` は `recurrence.seriesTime` プロパティを表す JSON のプレースホルダーです。 [`SeriesTime`][SeriesTime link] メソッドを使用して、繰り返しの日付と時刻のプロパティを取得する必要があります。

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

### <a name="get-the-recurrence-details"></a>定期的なアイテムの詳細を取得する

(`getAsync` コールバックまたは `item.recurrence` のいずれかから) 繰り返しオブジェクトを取得した後、特定の定期的なアイテムのプロパティを表示できます。 たとえば、プロパティの [methods][SeriesTime link] を使用して、系列の開始日と終了日時を `recurrence.seriesTime` 取得できます。

```js
// Get series date and time info
const seriesTime = recurrence.seriesTime;
const startTime = recurrence.seriesTime.getStartTime();
const endTime = recurrence.seriesTime.getEndTime();
const startDate = recurrence.seriesTime.getStartDate();
const endDate = recurrence.seriesTime.getEndDate();
const duration = recurrence.seriesTime.getDuration();

// Get series time zone
const timeZone = recurrence.recurrenceTimeZone;

// Get recurrence properties
const recurrenceProperties = recurrence.recurrenceProperties;

// Get recurrence type
const recurrenceType = recurrence.recurrenceType;
```

## <a name="see-also"></a>関連項目

- [RecurrenceChanged イベント](/javascript/api/office/office.eventtype)
- [Recurrence オブジェクト](/javascript/api/outlook/office.recurrence)
- [SeriesTime オブジェクト](/javascript/api/outlook/office.seriestime)
