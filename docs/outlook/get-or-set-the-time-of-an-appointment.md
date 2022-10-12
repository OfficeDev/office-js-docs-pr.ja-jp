---
title: Outlook アドインで予定の時刻を取得または設定する
description: Outlook アドインで予定の開始時間と終了時間を取得または設定する方法について説明します。
ms.date: 10/07/2022
ms.localizationpriority: medium
ms.openlocfilehash: c7aa40fda15c613aca869af8b277d4deb6fbf833
ms.sourcegitcommit: a2df9538b3deb32ae3060ecb09da15f5a3d6cb8d
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/12/2022
ms.locfileid: "68541234"
---
# <a name="get-or-set-the-time-when-composing-an-appointment-in-outlook"></a>Outlook で予定を作成するときに時刻を取得または設定する

Office JavaScript API は、ユーザーが作成している予定の開始時刻または終了時刻を取得および設定するための非同期メソッド ([Time.getAsync](/javascript/api/outlook/office.time#outlook-office-time-getasync-member(1)) と [Time.setAsync](/javascript/api/outlook/office.time#outlook-office-time-setasync-member(1))) を提供します。 これらの非同期メソッドは、アドインを作成する場合にのみ使用できます。これらのメソッドを使用するには、「作成フォーム用の Outlook アドインを作成する」の説明に従って、アドインの作成フォームをアクティブ化するために、Outlook に適したアドイン XML マニフェストを設定していることを確認 [します](compose-scenario.md)。 Office アドイン [の Teams マニフェスト (プレビュー) を](../develop/json-manifest-overview.md)使用するアドインでは、アクティブ化ルールはサポートされていません。

The [start](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) and [end](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) properties are available for appointments in both compose and read forms. In a read form, you can access the properties directly from the parent object, as in:

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

This section shows a code sample that gets the start time of the appointment that the user is composing and displays the time. You can use the same code and replace the **start** property by the **end** property to get the end time. This code sample assumes a rule in the add-in manifest that activates the add-in in a compose form for an appointment, as shown below.

```XML
<Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>
```

**item.start.getAsync** または **item.end.getAsync** を使用するには、非同期呼び出しの状態と結果をチェックするコールバック関数を指定します。 _asyncContext_ 省略可能なパラメーターを使用して、コールバック関数に必要な引数を指定できます。 You can obtain status, results and any error using the output parameter _asyncResult_ of the callback. If the asynchronous call is successful, you can get the start time as a **Date** object in UTC format using the [AsyncResult.value](/javascript/api/office/office.asyncresult#office-office-asyncresult-value-member) property.

```js
let item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready method.
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

This section shows a code sample that sets the start time of the appointment or message that the user is composing. You can use the same code and replace the **start** property by the **end** property to set the end time. Note that if the appointment compose form already has an existing start time, setting the start time subsequently will adjust the end time to maintain any previous duration for the appointment. If the appointment compose form already has an existing end time, setting the end time subsequently will adjust both the duration and end time. If the appointment has been set as an all-day event, setting the start time will adjust the end time to 24 hours later, and uncheck the UI for the all-day event in the compose form.

前の例と同様、このコード サンプルは、予定の新規作成フォームでアドインをアクティブ化するアドインのマニフェストのルールを想定しています。

**item.start.setAsync** または **item.end.setAsync** を使用するには、_dateTime_ パラメーターで UTC で **Date** 値を指定します。 クライアントでユーザーによる入力に基づいて日付を取得する場合は、[mailbox.convertToUtcClientTime](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) を使用して、値を UTC の **Date** オブジェクトに変換します。 _asyncContext_ パラメーターには、省略可能なコールバック関数とコールバック関数の引数を指定できます。 コールバックの _asyncResult_ 出力パラメーターで、状態、結果およびエラー メッセージを確認する必要があります。 非同期呼び出しが成功すると、指定した開始時刻または終了時刻文字列が **setAsync** によってプレーン テキストとして挿入され、そのアイテムの既存の開始時刻または終了時刻が上書きされます。

```js
let item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready method.
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
