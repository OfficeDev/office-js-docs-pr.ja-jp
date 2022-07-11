---
title: Outlook アドインで日付値を処理する
description: Office JavaScript API では、ほとんどの保存と日付と時刻の取得に JavaScript Date オブジェクトを使用します。
ms.date: 07/08/2022
ms.localizationpriority: medium
ms.openlocfilehash: 49de8db712400e006dc919e9ad62ae6cbaaa11cf
ms.sourcegitcommit: d8ea4b761f44d3227b7f2c73e52f0d2233bf22e2
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/11/2022
ms.locfileid: "66713078"
---
# <a name="tips-for-handling-date-values-in-outlook-add-ins"></a>Outlook アドインで日付値を処理するためのヒント

Office JavaScript API では、ほとんどの保存と日付と時刻の取得に JavaScript [Date](https://www.w3schools.com/jsref/jsref_obj_date.asp) オブジェクトを使用します。

この `Date` オブジェクトは、 [getUTCDate](https://www.w3schools.com/jsref/jsref_getutcdate.asp)、 [getUTCHour](https://www.w3schools.com/jsref/jsref_getutchours.asp)、 [getUTCMinutes](https://www.w3schools.com/jsref/jsref_getutcminutes.asp)、 [toUTCString](https://www.w3schools.com/jsref/jsref_toutcstring.asp) などのメソッドを提供します。このメソッドは、要求された日付または時刻の値を世界協定時刻 (UTC) 時刻に従って返します。

また、このオブジェクトには `Date` 、 [getDate](https://www.w3schools.com/jsref/jsref_getutcdate.asp)、 [getHour](https://www.w3schools.com/jsref/jsref_getutchours.asp)、 [getMinutes](https://www.w3schools.com/jsref/jsref_getminutes.asp)、 [toString](https://www.w3schools.com/jsref/jsref_tostring_date.asp) などの他のメソッドも用意されています。このメソッドは、"ローカル時刻" に従って要求された日付または時刻を返します。

"現地時刻" の概念は、主にクライアント コンピューター上のブラウザーおよびオペレーティング システムによって判断されます。 たとえば、Windows ベースのクライアント コンピューターで実行されているほとんどのブラウザーでは、JavaScript の呼び出しによって `getDate`、クライアント コンピューター上の Windows で設定されたタイム ゾーンに基づいて日付が返されます。

次の例では、ローカル時刻にオブジェクト`myLocalDate`を`Date`作成し、その日付を UTC の日付文字列に変換する呼び出し`toUTCString`を行います。

```js
// Create and get the current date represented 
// in the client computer time zone.
const myLocalDate = new Date (); 

// Convert the Date value in the client computer time zone
// to a date string in UTC, and display the string.
document.write ("The current UTC time is " + 
    myLocalDate.toUTCString());
```

JavaScript `Date` オブジェクトを使用して UTC またはクライアント コンピューターのタイム ゾーンに基づいて日付または時刻の値を取得できますが、 **Date** オブジェクトは一点で制限されます。他の特定のタイム ゾーンの日付または時刻の値を返すメソッドは提供されません。 たとえば、クライアント コンピューターが東部標準時 (EST) に設定されている場合、太平洋標準時 (PST) などの EST または UTC 以外の時間値を取得できるメソッドはありません `Date` 。

## <a name="date-related-features-for-outlook-add-ins"></a>Outlook アドインの日付関連機能

前述の JavaScript の制限は、Office JavaScript API を使用して、Outlook リッチ クライアントで実行される Outlook アドイン、およびOutlook on the webまたはモバイル デバイスで実行される Outlook アドインの日付または時刻の値を処理する場合に影響します。

### <a name="time-zones-for-outlook-clients"></a>Outlook クライアントのタイム ゾーン

わかりやすくするため、問題のタイム ゾーンを定義します。

|**タイム ゾーン**|**説明**|
|:-----|:-----|
|クライアント コンピューターのタイム ゾーン|これは、クライアント コンピューターのオペレーティング システムで設定されています。 ほとんどのブラウザーでは、クライアント コンピューターのタイム ゾーンを使用して、JavaScript `Date` オブジェクトの日付または時刻の値を表示します。<br/><br/>Outlook リッチ クライアントでは、このタイム ゾーンを使用して、ユーザー インターフェイスの日付または時刻の値を表示します。 <br/><br/>たとえば、Windows を実行しているクライアント コンピューター上の Outlook では、Windows 上で設定されているタイム ゾーンをローカル タイム ゾーンとして使用します。 Mac では、ユーザーがクライアント コンピューターのタイム ゾーンを変更した場合、Outlook on Mac でも Outlook でタイム ゾーンを更新するように求められます。|
|Exchange 管理センター (EAC) のタイム ゾーン|ユーザーは、初めてOutlook on the webまたはモバイル デバイスにログオンするときに、このタイム ゾーン値 (および優先言語) を設定します。 <br/><br/>Outlook on the webおよびモバイル デバイスでは、このタイム ゾーンを使用して、ユーザー インターフェイスに日付または時刻の値を表示します。|

Outlook リッチ クライアントではクライアント コンピューターのタイム ゾーンが使用され、Outlook on the webデバイスとモバイル デバイスのユーザー インターフェイスで EAC タイム ゾーンが使用されるため、同じメールボックスにインストールされている同じアドインのローカル時間は、Outlook リッチ クライアントとOutlook on the web またはモバイル デバイスで実行する場合に異なる場合があります。 Outlook アドイン開発者は日付の値を適切に入力して出力し、その値が、ユーザーから対応するクライアントに求められるタイム ゾーンと常に一致するようにする必要があります。

### <a name="date-related-api"></a>日付関連の API

日付関連の機能をサポートする Office JavaScript API のプロパティとメソッドを次に示します。

|API メンバー|タイム ゾーン表現|Outlook リッチ クライアントの例|Outlook on the webまたはモバイル デバイスの例|
|--------------|----------------------------|-------------------------------------|-------------------|
|[Office.context.mailbox.userProfile.timeZone](/javascript/api/outlook/office.userprofile?view=outlook-js-preview&preserve-view=true#outlook-office-userprofile-timezone-member)|Outlook リッチ クライアントでは、このプロパティはクライアント コンピューターのタイム ゾーンを返します。 Outlook on the webおよびモバイル デバイスでは、このプロパティは EAC タイム ゾーンを返します。 |EST|PST|
|[Office.context.mailbox.item.dateTimeCreated](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) および [Office.context.mailbox.item.dateTimeModified](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|これらの各プロパティは、JavaScript `Date` オブジェクトを返します。 次の例`myUTCDate`に示すように、この`Date`値は UTC で正しく、Outlook リッチ クライアント、Outlook on the web、モバイル デバイスでも同じ値を持ちます。<br/><br/>`const myDate = Office.mailbox.item.dateTimeCreated;`<br/>`const myUTCDate = myDate.getUTCDate;`<br/><br/>ただし、呼び出し`myDate.getDate`はクライアント コンピューターのタイム ゾーンで日付値を返します。これは、Outlook リッチ クライアント インターフェイスで日付時刻の値を表示するために使用されるタイム ゾーンと一致しますが、Outlook on the webおよびモバイル デバイスがユーザー インターフェイスで使用する EAC タイム ゾーンとは異なる場合があります。|アイテムが午前 9 時 (UTC) に作成される場合:<br/><br/>`Office.mailbox.item.`<br/>`dateTimeCreated.getHours` は、午前 4 時 (EST) を返します。<br/><br/>アイテムが午前 11 時 (UTC) に変更された場合:<br/><br/>`Office.mailbox.item.`<br/>`dateTimeModified.getHours` は、午前 6 時 (EST) を返します。|アイテムの作成時刻が午前 9 時 (UTC) の場合:<br/><br/>`Office.mailbox.item.`</br>`dateTimeCreated.getHours` は、午前 4 時 (EST) を返します。<br/><br/>アイテムが午前 11 時 (UTC) に変更された場合:<br/><br/>`Office.mailbox.item.`</br>`dateTimeModified.getHours` は、午前 6 時 (EST) を返します。<br/><br/>ユーザー インターフェイスで作成時刻や変更時刻を表示する場合は、まず時刻を PST に変換して、他のユーザー インターフェイスと一貫性を保つようにします。|
|[Office.context.mailbox.displayNewAppointmentForm](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)|_各開始_ パラメーターと _終了_ パラメーターには、JavaScript `Date` オブジェクトが必要です。 引数は、Outlook リッチ クライアント、Outlook on the web、モバイル デバイスのユーザー インターフェイスで使用されるタイム ゾーンに関係なく、UTC で正しく指定する必要があります。|予定フォームの開始時刻と終了時刻が 9 AM UTC と 11 AM UTC の場合は、引数と`end`引数が UTC で正しいことを確認`start`する必要があります。つまり、次のことを意味します。<br/><br/><ul><li>`start.getUTCHours` は午前 9 時 (UTC) を返します。</li><li>`end.getUTCHours` は午前 11 時 (UTC) を返します。</li></ul>|予定フォームの開始時刻と終了時刻が 9 AM UTC と 11 AM UTC の場合は、引数と`end`引数が UTC で正しいことを確認`start`する必要があります。つまり、次のことを意味します。<br/><br/><ul><li>`start.getUTCHours` は午前 9 時 (UTC) を返します。</li><li>`end.getUTCHours` は午前 11 時 (UTC) を返します。</li></ul>|

## <a name="helper-methods-for-date-related-scenarios"></a>日付関連のシナリオ向けのヘルパー メソッド

前のセクションで説明したように、Outlook on the webまたはモバイル デバイスのユーザーの "ローカル時刻" は Outlook リッチ クライアントでは異なる場合があるため、JavaScript **Date** オブジェクトはクライアント コンピューターのタイム ゾーンまたは UTC への変換のみをサポートしているため、Office JavaScript API には [、Office.context.mailbox.convertToLocalClientTime](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) と [Office.context.mailbox.convertToUtcClientTime](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) という 2 つのヘルパー メソッドが用意されています。

これらのヘルパー メソッドは、Outlook リッチ クライアント、Outlook on the web、モバイル デバイスでは、次の 2 つの日付関連のシナリオで日付または時刻を異なる方法で処理する必要があるため、アドインのさまざまなクライアントに対して "write-once" を強化します。

### <a name="scenario-a-displaying-item-creation-or-modified-time"></a>シナリオ A: アイテムの作成時刻または変更時刻の表示

項目の作成時刻 () または変更時刻 (`Item.dateTimeCreated``Item.dateTimeModified`ユーザー インターフェイス) を表示する場合は、まず`convertToLocalClientTime`、これらのプロパティによって提供されるオブジェクトを変換`Date`して、適切なローカル時間にディクショナリ表現を取得します。 その後、辞書の日付部分を表示します。 このシナリオの例を次に示します。

```js
// This date is UTC-correct.
const myDate = Office.context.mailbox.item.dateTimeCreated;

// Call helper method to get date in dictionary format, 
// represented in the appropriate local time.
// In an Outlook rich client, this is dictionary format 
// in client computer time zone.
// In Outlook on the web or mobile devices, this dictionary 
// format is in EAC time zone.
const myLocalDictionaryDate = Office.context.mailbox.convertToLocalClientTime(myDate);

// Display different parts of the dictionary date.
document.write ("The item was created at " + myLocalDictionaryDate["hours"] + 
    ":" + myLocalDictionaryDate["minutes"]);)
```

`convertToLocalClientTime` Outlook リッチ クライアントとOutlook on the webまたはモバイル デバイスの違いを処理します。

- 現在のアプリケーションがリッチ クライアントであることが検出された場合 `convertToLocalClientTime` 、このメソッドは、リッチ クライアント ユーザー インターフェイスの `Date` 残りの部分と一致する、同じクライアント コンピュータータイム ゾーン内のディクショナリ表現に表現を変換します。

- 現在のアプリケーションがOutlook on the webまたはモバイル デバイスであることが検出された場合`convertToLocalClientTime`、このメソッドは、UTC 正しい`Date`表現を EAC タイム ゾーンのディクショナリ形式に変換します。これは、他のOutlook on the webまたはモバイル デバイスのユーザー インターフェイスと一致します。

### <a name="scenario-b-displaying-start-and-end-dates-in-a-new-appointment-form"></a>シナリオ B: 新しい予定フォームの開始日付と終了日付の表示

ローカル時刻で表される日付と時刻の値の異なる部分を入力として取得し、このディクショナリ入力値を予定フォームの開始時刻または終了時刻として指定する場合は、まずヘルパー メソッドを使用 `convertToUtcClientTime` してディクショナリ値を UTC 正しい `Date` オブジェクトに変換します。

次の例では、`myLocalDictionaryStartDate` および `myLocalDictionaryEndDate` をユーザーから取得した辞書形式の日付と時刻の値と仮定しています。 これらの値は、クライアント プラットフォームに応じて、ローカル時刻に基づいています。

```js
const myUTCCorrectStartDate = Office.context.mailbox.convertToUtcClientTime(myLocalDictionaryStartDate);
const myUTCCorrectEndDate = Office.context.mailbox.convertToUtcClientTime(myLocalDictionaryEndDate);

```

出力結果の値 `myUTCCorrectStartDate` と `myUTCCorrectEndDate` は、正しい UTC です。 次に、これらの `Date` オブジェクトをメソッドの _Start_ パラメーターと _End_ パラメーターの `Mailbox.displayNewAppointmentForm` 引数として渡して、新しい予定フォームを表示します。

`convertToUtcClientTime` Outlook リッチ クライアントとOutlook on the webまたはモバイル デバイスの違いを処理します。

- 現在のアプリケーションが Outlook リッチ クライアントであることが検出された場合 `convertToUtcClientTime` 、このメソッドは単にディクショナリ表現をオブジェクトに `Date` 変換します。 この`Date`オブジェクトは UTC で正しく、.`displayNewAppointmentForm`

- 現在のアプリケーションがOutlook on the webまたはモバイル デバイスであることが検出された場合`convertToUtcClientTime`、このメソッドは、EAC タイム ゾーンで表される日付と時刻の値のディクショナリ形式を`Date`オブジェクトに変換します。 この`Date`オブジェクトは UTC で正しく、.`displayNewAppointmentForm`

## <a name="see-also"></a>関連項目

- [テスト用に Outlook アドインを展開してインストールする](testing-and-tips.md)
