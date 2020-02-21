---
title: Outlook アドインで日付値を処理する
description: JavaScript API for Office では、日付と時刻の保存および取得のほとんどの場合に、JavaScript の Date オブジェクトを使用します。
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: 5718839ebda433df6fb14886da34d734f81eb5f2
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/20/2020
ms.locfileid: "42166451"
---
# <a name="tips-for-handling-date-values-in-outlook-add-ins"></a>Outlook アドインで日付値を処理するためのヒント

JavaScript API for Office では、日付と時刻の保存および取得のほとんどで、JavaScript の [Date](https://www.w3schools.com/jsref/jsref_obj_date.asp) オブジェクトを使用します。 

その **Date** オブジェクトには、[getUTCDate](https://www.w3schools.com/jsref/jsref_getutcdate.asp)、[getUTCHour](https://www.w3schools.com/jsref/jsref_getutchours.asp)、[getUTCMinutes](https://www.w3schools.com/jsref/jsref_getutcminutes.asp)、および [toUTCString](https://www.w3schools.com/jsref/jsref_toutcstring.asp) などのメソッドが提供されており、それらは、要求された日付または時刻の値を、協定世界時 (UTC) 時刻に従って返します。

**Date** オブジェクトには、[getDate](https://www.w3schools.com/jsref/jsref_getutcdate.asp)、[getHour](https://www.w3schools.com/jsref/jsref_getutchours.asp)、[getMinutes](https://www.w3schools.com/jsref/jsref_getminutes.asp)、[toString](https://www.w3schools.com/jsref/jsref_tostring_date.asp) などのメソッドもあります。これらのメソッドは、要求された日付または時刻を "現地時刻" に従って返します。

"現地時刻" の概念は、主にクライアント コンピューター上のブラウザーおよびオペレーティング システムによって判断されます。たとえば、Windows ベースのクライアント コンピューター上で動作している大部分のブラウザーでは、JavaScript で **getDate** を呼び出すと、クライアント コンピューター上の Windows で設定されているタイム ゾーンに基づく日付が返されます。

次の例では、`myLocalDate` という名前の **Date** オブジェクトを現地時刻で作成し、**toUTCString** を呼び出して、その日付を UTC の日付文字列に変換します。

```js
// Create and get the current date represented 
// in the client computer time zone.
var myLocalDate = new Date (); 

// Convert the Date value in the client computer time zone
// to a date string in UTC, and display the string.
document.write ("The current UTC time is " + 
    myLocalDate.toUTCString());
```

JavaScript の **Date** オブジェクトを使用して UTC に基づく日付または時刻の値またはクライアント コンピュータのタイム ゾーンを取得できますが、**Date** オブジェクトは 1 つの点で制限があります。これには、他の特定のタイム ゾーンの日付または時刻の値を返すメソッドが用意されていないという点です。 たとえば、クライアント コンピュータが東部標準時 (EST) に設定されている場合、太平洋標準時 (PST) のように、EST でも UTC でもない時間値を取得するための **Date** メソッドはありません。


## <a name="date-related-features-for-outlook-add-ins"></a>Outlook アドインの日付関連機能

前述の JavaScript の制限は、outlook リッチクライアントおよび Outlook on the web またはモバイルデバイスで実行される Outlook アドインの日付または時刻の値を処理するために JavaScript API for Office を使用する場合に、意味があります。


### <a name="time-zones-for-outlook-clients"></a>Outlook クライアントのタイム ゾーン

わかりやすくするため、問題のタイム ゾーンを定義します。

|**タイム ゾーン**|**説明**|
|:-----|:-----|
|クライアント コンピューターのタイム ゾーン|これは、クライアント コンピューターのオペレーティング システムで設定されています。 ほとんどのブラウザーでは、クライアント コンピュータのタイム ゾーンを使用することにより、JavaScript の **Date** オブジェクトの日付または時刻の値を表示します。<br/><br/>Outlook リッチ クライアントでは、このタイム ゾーンを使用して、ユーザー インターフェイスの日付または時刻の値を表示します。 <br/><br/>たとえば、Windows を実行しているクライアント コンピューター上の Outlook では、Windows 上で設定されているタイム ゾーンをローカル タイム ゾーンとして使用します。 Mac では、ユーザーがクライアントコンピューターのタイムゾーンを変更した場合、Outlook on Mac はユーザーに Outlook のタイムゾーンの更新も要求します。|
|Exchange 管理センター (EAC) のタイム ゾーン|このタイムゾーンの値 (および優先する言語) は、ユーザーが初めて web 上の Outlook またはモバイルデバイスにログオンするときに設定されます。 <br/><br/>Web 上の Outlook およびモバイルデバイスは、このタイムゾーンを使用して、ユーザーインターフェイスに日付または時刻の値を表示します。|

Outlook リッチクライアントはクライアントコンピューターのタイムゾーンを使用しており、web およびモバイルデバイスの Outlook のユーザーインターフェイスは EAC タイムゾーンを使用しているため、同じメールボックスにインストールされている同じアドインのローカル時刻は、Outlook リッチクライアントで実行している場合と異なる場合があります。nt および Outlook on the web またはモバイルデバイス。 Outlook アドイン開発者は日付の値を適切に入力して出力し、その値が、ユーザーから対応するクライアントに求められるタイム ゾーンと常に一致するようにする必要があります。


### <a name="date-related-api"></a>日付関連の API

日付関連機能をサポートする JavaScript API for Office のプロパティおよびメソッドを以下に示します。

**API メンバー**|**タイム ゾーン表現**|**Outlook リッチ クライアントの例**|**Outlook on the web またはモバイルデバイスの例**
--------------|----------------------------|-------------------------------------|-------------------
[Office.context.mailbox.userProfile.timeZone](/javascript/api/outlook/office.userprofile?view=outlook-js-preview#timezone)|Outlook リッチ クライアントでは、このプロパティはクライアント コンピューターのタイム ゾーンを返します。 Outlook on the web およびモバイルデバイスの場合、このプロパティは EAC タイムゾーンを返します。 |EST|PST
[Office.context.mailbox.item.dateTimeCreated](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) および [Office.context.mailbox.item.dateTimeModified](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|これらのプロパティはどちらも、JavaScript の **Date** オブジェクトを返します。 この**日付**値は、次の例に示すように、UTC と`myUTCDate`同じです。これは、outlook リッチクライアント、web 上の outlook、およびモバイルデバイスでは同じ値を持つことになります。<br/><br/>`var myDate = Office.mailbox.item.dateTimeCreated;`<br/>`var myUTCDate = myDate.getUTCDate;`<br/><br/>ただし、呼び出し`myDate.getDate`では、クライアントコンピューターのタイムゾーンで日付値が返されます。これは、outlook リッチクライアントインターフェイスに日付時刻の値を表示するために使用されるタイムゾーンと一致しますが、web 上の outlook とモバイルデバイスがユーザーインターフェイスで使用する EAC タイムゾーンとは異なる場合があります。|アイテムが午前 9 時 (UTC) に作成された場合:<br/><br/>`Office.mailbox.item.`<br/>`dateTimeCreated.getHours` は、午前 4 時 (EST) を返します。<br/><br/>アイテムが午前 11 時 (UTC) に変更された場合:<br/><br/>`Office.mailbox.item.`<br/>`dateTimeModified.getHours` は、午前 6 時 (EST) を返します。|アイテムの作成時刻が午前 9 時 (UTC) の場合:<br/><br/>`Office.mailbox.item.`</br>`dateTimeCreated.getHours` は、午前 4 時 (EST) を返します。<br/><br/>アイテムが午前 11 時 (UTC) に変更された場合:<br/><br/>`Office.mailbox.item.`</br>`dateTimeModified.getHours` は、午前 6 時 (EST) を返します。<br/><br/>ユーザー インターフェイスで作成時刻や変更時刻を表示する場合は、まず時刻を PST に変換して、他のユーザー インターフェイスと一貫性を保つようにします。
[Office.context.mailbox.displayNewAppointmentForm](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)|_Start_ パラメーターと _End_ パラメーターには、それぞれ JavaScript の **Date** オブジェクトが必要です。 引数は、Outlook リッチクライアントのユーザーインターフェイスまたは web 上の Outlook またはモバイルデバイスで使用されているタイムゾーンに関係なく、正しい UTC である必要があります。|予定フォームの開始時刻と終了時刻が午前 9 時 (UTC) と午前 11 時 (UTC) の場合、`start` と `end` の引数は正しい UTC 時刻である必要があります。つまり、<br/><br/><ul><li>`start.getUTCHours` は午前 9 時 (UTC) を返します。</li><li>`end.getUTCHours` は午前 11 時 (UTC) を返します。</li></ul>|予定フォームの開始時刻と終了時刻が午前 9 時 (UTC) と午前 11 時 (UTC) の場合、`start` と `end` の引数は正しい UTC 時刻である必要があります。つまり、<br/><br/><ul><li>`start.getUTCHours` は午前 9 時 (UTC) を返します。</li><li>`end.getUTCHours` は午前 11 時 (UTC) を返します。</li></ul>

## <a name="helper-methods-for-date-related-scenarios"></a>日付関連のシナリオ向けのヘルパー メソッド


前のセクションで説明したように、web 上またはモバイルデバイス上の Outlook のユーザーの "ローカル時刻" は Outlook リッチクライアントでは異なる場合がありますが **、javascript API** for office では、クライアントコンピューターのタイムゾーンまたは UTC への変換をサポートして[いますが](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)、javascript API for Office では、 [convertToUtcClientTime](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)という2つのヘルパーメソッドが用意されています。

これらのヘルパーメソッドは、次の2つの日付関連のシナリオ、Outlook リッチクライアント、web およびモバイルデバイスで、日付または時刻を処理する必要があるかどうかを処理する必要があるため、アドインのさまざまなクライアントに対して "ライトワンス" を強化しています。


### <a name="scenario-a-displaying-item-creation-or-modified-time"></a>シナリオ A: アイテムの作成時刻または変更時刻の表示

ユーザー インターフェイスにアイテムの作成時刻 (**Item.dateTimeCreated**) または変更時刻 (**Item.dateTimeModified**) を表示している場合、まず **convertToLocalClientTime** を使用して、これらのプロパティで提供される **Date** オブジェクトを変換し、適切な現地時刻の辞書表現を取得します。 その後、辞書の日付部分を表示します。 このシナリオの例を次に示します。


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

**Converttolocalclienttime**は、outlook リッチクライアントと web 上の outlook、またはモバイルデバイスの違いに対処することに注意してください。


- **convertToLocalClientTime** メソッドでは、現在のホストがリッチ クライアントであることを検出すると、**Date** 表現を同じクライアント コンピューター タイム ゾーンの辞書表現に変換して、他のリッチ クライアント ユーザー インターフェイスとの一貫性を保ちます。
    
- **Converttolocalclienttime**が現在のホストを web またはモバイルデバイスで検出した場合、このメソッドは、UTC-正しい**日付**表現を EAC タイムゾーンの辞書形式に変換し、他の outlook on the web またはモバイルデバイスのユーザーインターフェイスとの一貫性を保ちます。
    

### <a name="scenario-b-displaying-start-and-end-dates-in-a-new-appointment-form"></a>シナリオ B: 新しい予定フォームの開始日付と終了日付の表示

現地時刻で表された日付と時刻の値の異なる各部分を、入力として取得しているときに、この辞書の入力値を予定フォームの開始時刻または終了時刻として提供する場合は、まず **convertToUtcClientTime** ヘルパー メソッドを使用して、ディクショナリ値を正しい UTC の **Date** オブジェクトに変換します。

次の例では、`myLocalDictionaryStartDate` および `myLocalDictionaryEndDate` をユーザーから取得した辞書形式の日付と時刻の値と仮定しています。これらの値は、ホスト アプリケーションに依存する、現地時刻に基づいています。

```js
var myUTCCorrectStartDate = Office.context.mailbox.convertToUtcClientTime(myLocalDictionaryStartDate);
var myUTCCorrectEndDate = Office.context.mailbox.convertToUtcClientTime(myLocalDictionaryEndDate);

```

出力結果の値 `myUTCCorrectStartDate` と `myUTCCorrectEndDate` は、正しい UTC です。 次に、これらの **Date** オブジェクトを **Mailbox.displayNewAppointmentForm** メソッドの _Start_ パラメーターと _End_ パラメーターの引数として渡し、新しい予定フォームを表示します。

**ConvertToUtcClientTime**は、outlook リッチクライアントと web 上の outlook またはモバイルデバイスとの違いに注意してください。


- **convertToUtcClientTime** により、現在のホストが Outlook リッチ クライアントであることが検出された場合、このメソッドは、単に辞書表現を **Date** オブジェクトに変換します。 この **Date** オブジェクトは、**displayNewAppointmentForm** で想定される正しい UTC です。
    
- **ConvertToUtcClientTime**は、現在のホストが web またはモバイルデバイス上にあることを検出すると、EAC のタイムゾーンで表現される日付と時刻の値のディクショナリ形式を**date**オブジェクトに変換します。 この **Date** オブジェクトは、**displayNewAppointmentForm** で想定される正しい UTC です。
    

## <a name="see-also"></a>関連項目

- [テスト用に Outlook アドインを展開してインストールする](testing-and-tips.md)
    


