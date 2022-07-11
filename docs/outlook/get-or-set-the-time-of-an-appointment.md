---
title: Outlook アドインで予定の時刻を取得または設定する
description: Outlook アドインで予定の開始時間と終了時間を取得または設定する方法について説明します。
ms.date: 07/08/2022
ms.localizationpriority: medium
ms.openlocfilehash: 9c98ae89c4c078e77a07724536498c7791db9d05
ms.sourcegitcommit: d8ea4b761f44d3227b7f2c73e52f0d2233bf22e2
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/11/2022
ms.locfileid: "66713050"
---
# <a name="get-or-set-the-time-when-composing-an-appointment-in-outlook"></a>Outlook で予定を作成するときに時刻を取得または設定する

Office JavaScript API は、ユーザーが作成している予定の開始時刻または終了時刻を取得および設定するための非同期メソッド ([Time.getAsync](/javascript/api/outlook/office.time#outlook-office-time-getasync-member(1)) と [Time.setAsync](/javascript/api/outlook/office.time#outlook-office-time-setasync-member(1))) を提供します。 これらの非同期メソッドは、アドインを作成する場合にのみ使用できます。これらのメソッドを使用するには、「作成フォーム用の Outlook アドインを作成する」の説明に従って、Outlook がアドイン作成フォームをアクティブ化するために適切 [にアドイン](compose-scenario.md) マニフェストを設定していることを確認します。

[start](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) プロパティおよび [end](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) プロパティは、新規作成フォームと閲覧フォームの両方の予定で利用できます。閲覧フォームでは親オブジェクトから直接プロパティにアクセスでき、それには次を使用します。

```js
item.start
```

、

```js
item.end
```

ただし新規作成フォームでは、ユーザーとアドインの両方が同時に時刻を挿入または変更できるため、次に示すように、非同期メソッド **getAsync** を使用して開始時刻または終了時刻を取得する必要があります。

```js
item.start.getAsync
```

、および

```js
item.end.getAsync
```

Office JavaScript API のほとんどの非同期メソッドと同様に、 **getAsync** と **setAsync** は省略可能な入力パラメーターを受け取ります。 これらのオプションの入力パラメーターの指定について詳しくは、「 [Office アドインにおける非同期プログラミング](../develop/asynchronous-programming-in-office-add-ins.md#pass-optional-parameters-inline)」の「 [オプションのパラメーターを非同期メソッドに渡す](../develop/asynchronous-programming-in-office-add-ins.md)」を参照してください。

## <a name="get-the-start-or-end-time"></a>開始時刻または終了時刻を取得する

このセクションでは、ユーザーが作成している予定の開始時刻を取得して、その時刻を表示するサンプル コードについて説明します。同じコードを使用して、**start** プロパティを **end** プロパティに置き換えると終了時刻を取得できます。このサンプル コードは、以下に示すアドイン マニフェストのルールによって予定の新規作成フォームでアドインがアクティブになることを想定しています。

```XML
<Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>
```

**item.start.getAsync** または **item.end.getAsync** を使用する場合は、非同期呼び出しの状態と結果を確認するコールバック メソッドを用意します。_asyncContext_ オプション パラメーターを使用して、コールバック メソッドに必要な引数を指定できます。コールバックの出力パラメーター _asyncResult_ を使用して、状態、結果およびエラーを取得できます。非同期呼び出しに成功すると、**AsyncResult.value** プロパティを使用して開始時刻を UTC 形式の [Date](/javascript/api/office/office.asyncresult#office-office-asyncresult-value-member) オブジェクトとして取得できます。

```js
let item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get the start time of the item being composed.
        getStartTime();
    });
}

// Get the start time of the item that the user is composing.
function getStartTime() {
    item.start.getAsync(
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully got the start time, display it, first in UTC and 
                // then convert the Date object to local time and display that.
                write ('The start time in UTC is: ' + asyncResult.value.toString());
                write ('The start time in local time is: ' + asyncResult.value.toLocaleString());
            }
        });
}

// Write to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

## <a name="set-the-start-or-end-time"></a>開始時刻または終了時刻を設定する

ここでは、ユーザーが作成している予定またはメッセージの開始時刻を設定するサンプル コードについて説明します。同じコードを使用して、**start** プロパティを **end** プロパティに置き換えると終了時刻を設定できます。予定の新規作成フォームに既存の開始時刻がある場合、後で開始時刻を設定すると、以前の予定の期間が保たれるように終了時刻を調整します。予定の新規作成フォームに既存の終了時刻がある場合、後で終了時刻を設定すると、期間と終了時刻の両方が調整されます。予定が終日イベントとして設定されている場合、開始時刻を設定すると、終了時刻を 24 時間後に調整し、新規作成フォームの終日イベントの UI をオフにします。

前の例と同様、このコード サンプルは、予定の新規作成フォームでアドインをアクティブ化するアドインのマニフェストのルールを想定しています。

**item.start.setAsync** または **item.end.setAsync** を使用する場合は、**dateTime** パラメーターに UTC で _Date_ の値を指定します。クライアントでユーザーによる入力に基づいて日付を取得する場合は、[mailbox.convertToUtcClientTime](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) を使用して、値を UTC の **Date** オブジェクトに変換します。オプションのコールバック メソッドと、_asyncContext_ パラメーターでそのコールバック メソッドの引数を指定できます。コールバックの _asyncResult_ 出力パラメーターで、状態、結果およびエラー メッセージを確認する必要があります。非同期呼び出しが成功すると、指定した開始時刻または終了時刻文字列が **setAsync** によってプレーン テキストとして挿入され、そのアイテムの既存の開始時刻または終了時刻が上書きされます。

```js
let item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set the start time of the item being composed.
        setStartTime();
    });
}

// Set the start time of the item that the user is composing.
function setStartTime() {
    const startDate = new Date("September 27, 2012 12:30:00");
    
    item.start.setAsync(
        startDate,
        { asyncContext: { var1: 1, var2: 2 } },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully set the start time.
                // Do whatever appropriate for your scenario
                // using the arguments var1 and var2 as applicable.
            }
        });
}

// Write to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

## <a name="see-also"></a>関連項目

- [Outlook で新規作成フォームのアイテム データを取得および設定する](get-and-set-item-data-in-a-compose-form.md)
- [閲覧または新規作成フォームの Outlook アイテム データを取得および設定する](item-data.md)
- [新規作成フォーム用の Outlook アドインを作成する](compose-scenario.md)
- [Office アドインにおける非同期プログラミング](../develop/asynchronous-programming-in-office-add-ins.md)
- [Outlook の予定またはメッセージを作成するときに受信者を取得、設定、追加する](get-set-or-add-recipients.md)  
- [Outlook で予定またはメッセージを作成するときに件名を取得または設定する](get-or-set-the-subject.md)
- [Outlook で予定またはメッセージを作成するときに本文にデータを挿入する](insert-data-in-the-body.md)
- [Outlook で予定を作成するときに場所を取得または設定する](get-or-set-the-location-of-an-appointment.md)
