---
title: Outlook アドインで日付値を処理する
description: Office JavaScript API は、日付と時刻の保存および取得のほとんどに、JavaScript の Date オブジェクトを使用します。
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: fb27e7393da9f5192daa5f7b14099f3fb0aeded0
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/11/2020
ms.locfileid: "47431046"
---
# <a name="tips-for-handling-date-values-in-outlook-add-ins"></a>Outlook アドインで日付値を処理するためのヒント

Office JavaScript API は、日付と時刻の保存および取得のほとんどに、JavaScript の [Date](https://www.w3schools.com/jsref/jsref_obj_date.asp) オブジェクトを使用します。 

`Date`このオブジェクトは、要求された日付または時刻の値を世界協定時刻 (UTC) の時刻に基づいて返す[Getutcdate](https://www.w3schools.com/jsref/jsref_getutcdate.asp)、 [getUTCHour](https://www.w3schools.com/jsref/jsref_getutchours.asp)、 [Getutcdate](https://www.w3schools.com/jsref/jsref_getutcminutes.asp)、 [toutcstring](https://www.w3schools.com/jsref/jsref_toutcstring.asp)などのメソッドを提供します。

このオブジェクトには、 `Date` [getDate](https://www.w3schools.com/jsref/jsref_getutcdate.asp)、 [getHour](https://www.w3schools.com/jsref/jsref_getutchours.asp)、 [getminutes](https://www.w3schools.com/jsref/jsref_getminutes.asp)、 [toString](https://www.w3schools.com/jsref/jsref_tostring_date.asp)などの他のメソッドもあります。これは、要求された日付または時刻を "現地時刻" に従って返します。

"現地時刻" の概念は、主にクライアント コンピューター上のブラウザーおよびオペレーティング システムによって判断されます。 たとえば、Windows ベースのクライアントコンピューターで実行されているほとんどのブラウザーでは、への JavaScript 呼び出しは `getDate` 、クライアントコンピューターの Windows で設定されているタイムゾーンに基づいて日付を返します。

次の例では `Date` 、 `myLocalDate` ローカル時刻でオブジェクトを作成し、 `toUTCString` その日付を UTC の日付文字列に変換するための呼び出しを実行します。

```js
// Create and get the current date represented 
// in the client computer time zone.
var myLocalDate = new Date (); 

// Convert the Date value in the client computer time zone
// to a date string in UTC, and display the string.
document.write ("The current UTC time is " + 
    myLocalDate.toUTCString());
```

JavaScript オブジェクトを使用して、 `Date` UTC またはクライアントコンピューターのタイムゾーンに基づいて日付または時刻の値を取得できますが、 **date** オブジェクトは1つの尊敬で制限されていますが、その他の特定のタイムゾーンの日付または時刻の値を返すメソッドは提供しません。 たとえば、クライアントコンピューターが東部標準時 (EST) に設定されている場合、 `Date` 太平洋標準時 (PST) のように、est または UTC 以外の時間の値を取得するための方法はありません。


## <a name="date-related-features-for-outlook-add-ins"></a>Outlook アドインの日付関連機能

前述の JavaScript 制限は、Office JavaScript API を使用して outlook リッチクライアントで実行される Outlook アドインの日付または時刻の値を処理するとき、および Outlook on the web またはモバイルデバイスで使用する場合に、ユーザーにとって意味があります。


### <a name="time-zones-for-outlook-clients"></a>Outlook クライアントのタイム ゾーン

わかりやすくするため、問題のタイム ゾーンを定義します。

|**タイム ゾーン**|**説明**|
|:-----|:-----|
|クライアント コンピューターのタイム ゾーン|これは、クライアント コンピューターのオペレーティング システムで設定されています。 ほとんどのブラウザーでは、クライアントコンピューターのタイムゾーンを使用して、JavaScript オブジェクトの日付または時刻の値を表示し `Date` ます。<br/><br/>Outlook リッチ クライアントでは、このタイム ゾーンを使用して、ユーザー インターフェイスの日付または時刻の値を表示します。 <br/><br/>たとえば、Windows を実行しているクライアント コンピューター上の Outlook では、Windows 上で設定されているタイム ゾーンをローカル タイム ゾーンとして使用します。 Mac では、ユーザーがクライアントコンピューターのタイムゾーンを変更した場合、Outlook on Mac はユーザーに Outlook のタイムゾーンの更新も要求します。|
|Exchange 管理センター (EAC) のタイム ゾーン|このタイムゾーンの値 (および優先する言語) は、ユーザーが初めて web 上の Outlook またはモバイルデバイスにログオンするときに設定されます。 <br/><br/>Web 上の Outlook およびモバイルデバイスは、このタイムゾーンを使用して、ユーザーインターフェイスに日付または時刻の値を表示します。|

Outlook リッチクライアントではクライアントコンピューターのタイムゾーンが使用され、web およびモバイルデバイスのユーザーインターフェイスでは EAC タイムゾーンが使用されているため、同じメールボックスにインストールされている同じアドインのローカル時刻は、Outlook リッチクライアントで実行している場合と web またはモバイルデバイスで実行する場合とで異なる場合があります。 Outlook アドイン開発者は日付の値を適切に入力して出力し、その値が、ユーザーから対応するクライアントに求められるタイム ゾーンと常に一致するようにする必要があります。


### <a name="date-related-api"></a>日付関連の API

以下に、日付関連機能をサポートする Office JavaScript API のプロパティとメソッドを示します。

**API メンバー**|**タイム ゾーン表現**|**Outlook リッチ クライアントの例**|**Outlook on the web またはモバイルデバイスの例**
--------------|----------------------------|-------------------------------------|-------------------
[Office.context.mailbox.userProfile.timeZone](/javascript/api/outlook/office.userprofile?view=outlook-js-preview&preserve-view=true#timezone)|Outlook リッチ クライアントでは、このプロパティはクライアント コンピューターのタイム ゾーンを返します。 Outlook on the web およびモバイルデバイスの場合、このプロパティは EAC タイムゾーンを返します。 |EST|PST
[Office.context.mailbox.item.dateTimeCreated](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) および [Office.context.mailbox.item.dateTimeModified](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|これらの各プロパティは、JavaScript `Date` オブジェクトを返します。 `Date`次の例に示すように、この値は UTC-正しいです。これは、 `myUTCDate` outlook リッチクライアント、web 上の outlook、およびモバイルデバイスで同じ値を持つことになります。<br/><br/>`var myDate = Office.mailbox.item.dateTimeCreated;`<br/>`var myUTCDate = myDate.getUTCDate;`<br/><br/>ただし、呼び出しでは、  `myDate.getDate` クライアントコンピューターのタイムゾーンで日付値が返されます。これは、outlook リッチクライアントインターフェイスに日付時刻の値を表示するために使用されるタイムゾーンと一致しますが、web 上の outlook とモバイルデバイスがユーザーインターフェイスで使用する EAC タイムゾーンとは異なる場合があります。|アイテムが午前 9 時 (UTC) に作成された場合:<br/><br/>`Office.mailbox.item.`<br/>`dateTimeCreated.getHours` は、午前 4 時 (EST) を返します。<br/><br/>アイテムが午前 11 時 (UTC) に変更された場合:<br/><br/>`Office.mailbox.item.`<br/>`dateTimeModified.getHours` は、午前 6 時 (EST) を返します。|アイテムの作成時刻が午前 9 時 (UTC) の場合:<br/><br/>`Office.mailbox.item.`</br>`dateTimeCreated.getHours` は、午前 4 時 (EST) を返します。<br/><br/>アイテムが午前 11 時 (UTC) に変更された場合:<br/><br/>`Office.mailbox.item.`</br>`dateTimeModified.getHours` は、午前 6 時 (EST) を返します。<br/><br/>ユーザー インターフェイスで作成時刻や変更時刻を表示する場合は、まず時刻を PST に変換して、他のユーザー インターフェイスと一貫性を保つようにします。
[Office.context.mailbox.displayNewAppointmentForm](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)|_Start_および_End_パラメーターには、JavaScript オブジェクトが必要です `Date` 。 引数は、Outlook リッチクライアントのユーザーインターフェイスまたは web 上の Outlook またはモバイルデバイスで使用されているタイムゾーンに関係なく、正しい UTC である必要があります。|予定フォームの開始時刻と終了時刻が午前 9 時 (UTC) と午前 11 時 (UTC) の場合、`start` と `end` の引数は正しい UTC 時刻である必要があります。つまり、<br/><br/><ul><li>`start.getUTCHours` は午前 9 時 (UTC) を返します。</li><li>`end.getUTCHours` は午前 11 時 (UTC) を返します。</li></ul>|予定フォームの開始時刻と終了時刻が午前 9 時 (UTC) と午前 11 時 (UTC) の場合、`start` と `end` の引数は正しい UTC 時刻である必要があります。つまり、<br/><br/><ul><li>`start.getUTCHours` は午前 9 時 (UTC) を返します。</li><li>`end.getUTCHours` は午前 11 時 (UTC) を返します。</li></ul>

## <a name="helper-methods-for-date-related-scenarios"></a>日付関連のシナリオ向けのヘルパー メソッド


前のセクションで説明したように、web 上またはモバイルデバイス上の Outlook のユーザーの "現地時刻" は、Outlook リッチクライアントでは異なる場合がありますが、JavaScript の**Date**オブジェクトはクライアントコンピューターのタイムゾーンまたは UTC への変換をサポートしていますが、OFFICE javascript API には、2つのヘルパーメソッド[convertToUtcClientTime](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)が用意され[ています](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)。

これらのヘルパーメソッドは、次の2つの日付関連のシナリオ、Outlook リッチクライアント、web およびモバイルデバイスで、日付または時刻を処理する必要があるかどうかを処理する必要があるため、アドインのさまざまなクライアントに対して "ライトワンス" を強化しています。


### <a name="scenario-a-displaying-item-creation-or-modified-time"></a>シナリオ A: アイテムの作成時刻または変更時刻の表示

アイテムの作成時刻 () または変更時刻を表示している場合 `Item.dateTimeCreated` ( `Item.dateTimeModified` ユーザーインターフェイスでは、最初 `convertToLocalClientTime` に、これらの `Date` プロパティによって提供されるオブジェクトを変換して、適切なローカル時刻で辞書表現を取得します。 その後、辞書の日付部分を表示します。 このシナリオの例を次に示します。


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

`convertToLocalClientTime`Outlook リッチクライアントと web 上の outlook またはモバイルデバイスとの違いに注意してください。


- 現在のアプリケーションがリッチクライアントであることが検出された場合、この `convertToLocalClientTime` メソッドは `Date` 表現を同じクライアントコンピューターのタイムゾーンの辞書表現に変換し、その他のリッチクライアントユーザーインターフェイスとの一貫性を保ちます。
    
- `convertToLocalClientTime`現在のアプリケーションが web またはモバイルデバイス上にあることが検出された場合、このメソッドは、UTC-正しい `Date` 表現を EAC タイムゾーンの辞書形式に変換します。これは、web 上の outlook またはモバイルデバイスのユーザーインターフェイスと一致します。
    

### <a name="scenario-b-displaying-start-and-end-dates-in-a-new-appointment-form"></a>シナリオ B: 新しい予定フォームの開始日付と終了日付の表示

現地時刻で表された日付と時刻の値のさまざまな部分を入力として取得していて、この辞書入力値を予定フォームの開始時刻または終了時刻として提供したい場合は、最初に、ヘルパーメソッドを使用して、 `convertToUtcClientTime` 辞書値を正しいオブジェクトに変換し `Date` ます。

次の例では、`myLocalDictionaryStartDate` および `myLocalDictionaryEndDate` をユーザーから取得した辞書形式の日付と時刻の値と仮定しています。 これらの値は、クライアントプラットフォームに応じてローカル時刻に基づいています。

```js
var myUTCCorrectStartDate = Office.context.mailbox.convertToUtcClientTime(myLocalDictionaryStartDate);
var myUTCCorrectEndDate = Office.context.mailbox.convertToUtcClientTime(myLocalDictionaryEndDate);

```

出力結果の値 `myUTCCorrectStartDate` と `myUTCCorrectEndDate` は、正しい UTC です。 次に、 `Date` メソッドの _Start_ および _End_ パラメーターとして、これらのオブジェクトを引数として渡して、 `Mailbox.displayNewAppointmentForm` 新しい予定のフォームを表示します。

`convertToUtcClientTime`Outlook リッチクライアントと web 上の outlook またはモバイルデバイスとの違いに注意してください。


- `convertToUtcClientTime`現在のアプリケーションが Outlook リッチクライアントであることが検出された場合、このメソッドは単に辞書表現をオブジェクトに変換し `Date` ます。 この `Date` オブジェクトは、必要に応じて、UTC-正しい `displayNewAppointmentForm` 。
    
- `convertToUtcClientTime`現在のアプリケーションが web またはモバイルデバイス上にあることが検出された場合、メソッドは EAC タイムゾーンで表される日付と時刻の値のディクショナリ形式をオブジェクトに変換し `Date` ます。 この `Date` オブジェクトは、必要に応じて、UTC-正しい `displayNewAppointmentForm` 。
    
## <a name="see-also"></a>関連項目

- [テスト用に Outlook アドインを展開してインストールする](testing-and-tips.md)
