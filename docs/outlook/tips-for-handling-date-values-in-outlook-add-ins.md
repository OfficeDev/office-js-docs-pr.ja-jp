---
title: Outlook アドインで日付値を処理する
description: JavaScript API Office JavaScript Date オブジェクトは、ほとんどの日付と時刻の格納および取得に使用されます。
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: 4f0e93d147760f91c55fd5666f02b2c6cc5d5470
ms.sourcegitcommit: 7482ab6bc258d98acb9ba9b35c7dd3b5cc5bed21
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/24/2021
ms.locfileid: "51177993"
---
# <a name="tips-for-handling-date-values-in-outlook-add-ins"></a>Outlook アドインで日付値を処理するためのヒント

JavaScript API Office JavaScript [Date](https://www.w3schools.com/jsref/jsref_obj_date.asp) オブジェクトは、ほとんどのストレージと日付と時刻の取得に使用されます。 

`Date`このオブジェクトは、getUTCDate、getUTCHour、getUTCMinutes、toUTCString[](https://www.w3schools.com/jsref/jsref_toutcstring.asp)[](https://www.w3schools.com/jsref/jsref_getutcdate.asp)[](https://www.w3schools.com/jsref/jsref_getutchours.asp)などのメソッドを提供し、要求された日付または時刻の値をユニバーサル座標時刻 (UTC)[](https://www.w3schools.com/jsref/jsref_getutcminutes.asp)時刻に従って返します。

オブジェクト `Date` には、getDate、getHour、getMinutes、toString[](https://www.w3schools.com/jsref/jsref_tostring_date.asp)[](https://www.w3schools.com/jsref/jsref_getminutes.asp)[](https://www.w3schools.com/jsref/jsref_getutcdate.asp)などの他のメソッドも用意されています。このメソッドは、要求された日付または時刻を "ローカル時刻" に従って返します。 [](https://www.w3schools.com/jsref/jsref_getutchours.asp)

"現地時刻" の概念は、主にクライアント コンピューター上のブラウザーおよびオペレーティング システムによって判断されます。 たとえば、Windows ベースのクライアント コンピューターで実行されているほとんどのブラウザーでは、JavaScript の呼び出しは、クライアント コンピューター上の Windows で設定されたタイム ゾーンに基づいて日付 `getDate` を返します。

次の例では、ローカル時間にオブジェクトを作成し、その日付を UTC の日付文字列に変換 `Date` `myLocalDate` `toUTCString` する呼び出しを行います。

```js
// Create and get the current date represented 
// in the client computer time zone.
var myLocalDate = new Date (); 

// Convert the Date value in the client computer time zone
// to a date string in UTC, and display the string.
document.write ("The current UTC time is " + 
    myLocalDate.toUTCString());
```

JavaScript オブジェクトを使用して、UTC またはクライアント コンピューターのタイム ゾーンに基づいて日付または時刻の値を取得することができますが `Date` **、Date** オブジェクトは 1 つの点で制限されます。他の特定のタイム ゾーンの日付または時刻の値を返すメソッドは提供されません。 たとえば、クライアント コンピューターが東部標準時 (EST) に設定されている場合、太平洋標準時 (PST) などの EST または UTC 以外の時間値を取得できるメソッドはありません `Date` 。


## <a name="date-related-features-for-outlook-add-ins"></a>Outlook アドインの日付関連機能

前述の JavaScript の制限は、Office JavaScript API を使用して Outlook リッチ クライアントで実行される Outlook アドイン、および Outlook on the web またはモバイル デバイスで日付または時刻の値を処理する場合に、影響を及ぼす可能性があります。


### <a name="time-zones-for-outlook-clients"></a>Outlook クライアントのタイム ゾーン

わかりやすくするため、問題のタイム ゾーンを定義します。

|**タイム ゾーン**|**説明**|
|:-----|:-----|
|クライアント コンピューターのタイム ゾーン|これは、クライアント コンピューターのオペレーティング システムで設定されています。 ほとんどのブラウザーでは、クライアント コンピューターのタイム ゾーンを使用して、JavaScript オブジェクトの日付または時刻の値を表示 `Date` します。<br/><br/>Outlook リッチ クライアントでは、このタイム ゾーンを使用して、ユーザー インターフェイスの日付または時刻の値を表示します。 <br/><br/>たとえば、Windows を実行しているクライアント コンピューター上の Outlook では、Windows 上で設定されているタイム ゾーンをローカル タイム ゾーンとして使用します。 Mac では、ユーザーがクライアント コンピューターのタイム ゾーンを変更した場合、Outlook on Mac は Outlook でもタイム ゾーンを更新するようユーザーに求めるメッセージを表示します。|
|Exchange 管理センター (EAC) のタイム ゾーン|ユーザーが初めて Outlook on the web またはモバイル デバイスにログオンするときに、このタイム ゾーン値 (および優先言語) を設定します。 <br/><br/>Outlook on the web とモバイル デバイスでは、このタイム ゾーンを使用して、ユーザー インターフェイスに日付または時刻の値を表示します。|

Outlook リッチ クライアントはクライアント コンピューターのタイム ゾーンを使用し、Web およびモバイル デバイス上の Outlook のユーザー インターフェイスは EAC タイム ゾーンを使用します。Outlook リッチ クライアントと Outlook on the web またはモバイル デバイスで実行する場合、同じメールボックスにインストールされている同じアドインのローカル時間が異なる場合があります。 Outlook アドイン開発者は日付の値を適切に入力して出力し、その値が、ユーザーから対応するクライアントに求められるタイム ゾーンと常に一致するようにする必要があります。


### <a name="date-related-api"></a>日付関連の API

日付関連の機能をサポートする JavaScript API Officeプロパティとメソッドを次に示します。

|API メンバー|タイム ゾーン表現|Outlook リッチ クライアントの例|Outlook on the Web またはモバイル デバイスの例|
|--------------|----------------------------|-------------------------------------|-------------------|
|[Office.context.mailbox.userProfile.timeZone](/javascript/api/outlook/office.userprofile?view=outlook-js-preview&preserve-view=true#timezone)|Outlook リッチ クライアントでは、このプロパティはクライアント コンピューターのタイム ゾーンを返します。 Outlook on the web およびモバイル デバイスでは、このプロパティは EAC タイム ゾーンを返します。 |EST|PST|
|[Office.context.mailbox.item.dateTimeCreated](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) および [Office.context.mailbox.item.dateTimeModified](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|これらの各プロパティは、JavaScript オブジェクトを返 `Date` します。 次の例に示すように、この値は UTC が正しく、Outlook リッチ クライアント、Outlook on the web、モバイル デバイスで同じ `Date` `myUTCDate` 値を持っています。<br/><br/>`var myDate = Office.mailbox.item.dateTimeCreated;`<br/>`var myUTCDate = myDate.getUTCDate;`<br/><br/>ただし、呼び出しはクライアント コンピューターのタイム ゾーンの日付値を返します。これは、Outlook リッチ クライアント インターフェイスで日付時刻の値を表示するために使用されるタイム ゾーンと一致しますが、Outlook on the web およびモバイル デバイスがユーザー インターフェイスで使用する EAC タイム ゾーンとは異なる場合があります。 `myDate.getDate`|アイテムが午前 9 時 (UTC) に作成された場合:<br/><br/>`Office.mailbox.item.`<br/>`dateTimeCreated.getHours` は、午前 4 時 (EST) を返します。<br/><br/>アイテムが午前 11 時 (UTC) に変更された場合:<br/><br/>`Office.mailbox.item.`<br/>`dateTimeModified.getHours` は、午前 6 時 (EST) を返します。|アイテムの作成時刻が午前 9 時 (UTC) の場合:<br/><br/>`Office.mailbox.item.`</br>`dateTimeCreated.getHours` は、午前 4 時 (EST) を返します。<br/><br/>アイテムが午前 11 時 (UTC) に変更された場合:<br/><br/>`Office.mailbox.item.`</br>`dateTimeModified.getHours` は、午前 6 時 (EST) を返します。<br/><br/>ユーザー インターフェイスで作成時刻や変更時刻を表示する場合は、まず時刻を PST に変換して、他のユーザー インターフェイスと一貫性を保つようにします。|
|[Office.context.mailbox.displayNewAppointmentForm](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)|各 Start パラメーター  _と End_ パラメーター _には_ JavaScript オブジェクトが必要 `Date` です。 引数は、Outlook リッチ クライアントのユーザー インターフェイス、または Outlook on the web またはモバイル デバイスで使用されるタイム ゾーンに関係なく、UTC で正しい値を指定する必要があります。|予定フォームの開始時刻と終了時刻が午前 9 時 (UTC) と午前 11 時 (UTC) の場合、`start` と `end` の引数は正しい UTC 時刻である必要があります。つまり、<br/><br/><ul><li>`start.getUTCHours` は午前 9 時 (UTC) を返します。</li><li>`end.getUTCHours` は午前 11 時 (UTC) を返します。</li></ul>|予定フォームの開始時刻と終了時刻が午前 9 時 (UTC) と午前 11 時 (UTC) の場合、`start` と `end` の引数は正しい UTC 時刻である必要があります。つまり、<br/><br/><ul><li>`start.getUTCHours` は午前 9 時 (UTC) を返します。</li><li>`end.getUTCHours` は午前 11 時 (UTC) を返します。</li></ul>|

## <a name="helper-methods-for-date-related-scenarios"></a>日付関連のシナリオ向けのヘルパー メソッド


前のセクションで説明したように、Outlook on the web またはモバイル デバイスのユーザーの "現地時間" は Outlook リッチ クライアントでは異なる場合がありますが、JavaScript **Date** オブジェクトはクライアント コンピューターのタイム ゾーンまたは UTC への変換のみをサポートしています。Office JavaScript API には [、Office.context.mailbox.convertToLocalClientTime](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) と [Office.context.mailbox.convertToUtcTime](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)の 2 つのヘルパー メソッドが提供されています。

これらのヘルパー メソッドは、Outlook リッチ クライアント、Outlook on the web、およびモバイル デバイスで、次の 2 つの日付関連のシナリオで日付または時刻を異なる方法で処理する必要を処理するため、アドインの異なるクライアントに対して "write-once" を強化します。


### <a name="scenario-a-displaying-item-creation-or-modified-time"></a>シナリオ A: アイテムの作成時刻または変更時刻の表示

ユーザー インターフェイスでアイテムの作成時間 ( ) または変更時刻 ( ) を表示する場合は、まず、これらのプロパティによって提供されるオブジェクトを変換して、適切な現地時間で辞書表現を `Item.dateTimeCreated` `Item.dateTimeModified` 取得します `convertToLocalClientTime` `Date` 。 その後、辞書の日付部分を表示します。 このシナリオの例を次に示します。


```js
// This date is UTC-correct.
var myDate = Office.context.mailbox.item.dateTimeCreated;

// Call helper method to get date in dictionary format, 
// represented in the appropriate local time.
// In an Outlook rich client, this is dictionary format 
// in client computer time zone.
// In Outlook on the web or mobile devices, this dictionary 
// format is in EAC time zone.
var myLocalDictionaryDate = Office.context.mailbox.convertToLocalClientTime(myDate);

// Display different parts of the dictionary date.
document.write ("The item was created at " + myLocalDictionaryDate["hours"] + 
    ":" + myLocalDictionaryDate["minutes"]);)
```

Outlook リッチ クライアントと Outlook on the web またはモバイル デバイスの違いを注意 `convertToLocalClientTime` してください。


- 現在のアプリケーションがリッチ クライアントである場合、このメソッドは、リッチ クライアント ユーザー インターフェイスの残りの部分と一致する、同じクライアント コンピュータータイム ゾーン内の辞書表現に表現を `convertToLocalClientTime` `Date` 変換します。
    
- 現在のアプリケーションが Outlook on the Web またはモバイル デバイスである場合、このメソッドは UTC 正しい表現を EAC タイム ゾーンの辞書形式に変換し、Outlook on the web またはモバイル デバイスの他のユーザー インターフェイスと一致します。 `convertToLocalClientTime` `Date`
    

### <a name="scenario-b-displaying-start-and-end-dates-in-a-new-appointment-form"></a>シナリオ B: 新しい予定フォームの開始日付と終了日付の表示

現地時間で表される日付と時刻の値の異なる部分を入力として取得し、この辞書入力値を予定フォームの開始時刻または終了時刻として指定する場合は、まずヘルパー メソッドを使用して辞書値を UTC 正しいオブジェクトに変換します。 `convertToUtcClientTime` `Date`

次の例では、`myLocalDictionaryStartDate` および `myLocalDictionaryEndDate` をユーザーから取得した辞書形式の日付と時刻の値と仮定しています。 これらの値は、クライアント プラットフォームに依存する現地時間に基づいて設定されます。

```js
var myUTCCorrectStartDate = Office.context.mailbox.convertToUtcClientTime(myLocalDictionaryStartDate);
var myUTCCorrectEndDate = Office.context.mailbox.convertToUtcClientTime(myLocalDictionaryEndDate);

```

出力結果の値 `myUTCCorrectStartDate` と `myUTCCorrectEndDate` は、正しい UTC です。 次に、これらのオブジェクトをメソッドの Start パラメーターと End パラメーターの引数として渡して、新 `Date` しい予定 `Mailbox.displayNewAppointmentForm` フォームを表示します。

Outlook リッチ クライアントと Outlook on the web またはモバイル デバイスの違いを注意 `convertToUtcClientTime` してください。


- 現在のアプリケーションが Outlook リッチ クライアントである場合、メソッドは単に辞書表現を `convertToUtcClientTime` オブジェクトに変換 `Date` します。 この `Date` オブジェクトは UTC で正しく、予期される値です `displayNewAppointmentForm` 。
    
- 現在のアプリケーションが Outlook on the Web またはモバイル デバイスである場合、このメソッドは EAC タイム ゾーンで表される日付と時刻の値の辞書形式をオブジェクトに `convertToUtcClientTime` 変換 `Date` します。 この `Date` オブジェクトは UTC で正しく、予期される値です `displayNewAppointmentForm` 。
    
## <a name="see-also"></a>関連項目

- [テスト用に Outlook アドインを展開してインストールする](testing-and-tips.md)
