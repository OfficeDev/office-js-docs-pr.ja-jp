---
title: Outlook アドインで予定の時刻を取得または設定する
description: Outlook アドインで予定の開始時間と終了時間を取得または設定する方法について説明します。
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: d07d461b852e523626946a79a5c9c5e21c95fcdc
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324962"
---
# <a name="get-or-set-the-time-when-composing-an-appointment-in-outlook"></a><span data-ttu-id="a1b68-103">Outlook で予定を作成するときに時刻を取得または設定する</span><span class="sxs-lookup"><span data-stu-id="a1b68-103">Get or set the time when composing an appointment in Outlook</span></span>

<span data-ttu-id="a1b68-104">Office JavaScript API は、非同期メソッド ([getAsync](/javascript/api/outlook/office.Time#getasync-options--callback-)および[time async](/javascript/api/outlook/office.Time#setasync-datetime--options--callback-)) を使用して、ユーザーが作成している予定の開始時刻または終了時刻を取得および設定します。</span><span class="sxs-lookup"><span data-stu-id="a1b68-104">The Office JavaScript API provides asynchronous methods ([Time.getAsync](/javascript/api/outlook/office.Time#getasync-options--callback-) and [Time.setAsync](/javascript/api/outlook/office.Time#setasync-datetime--options--callback-)) to get and set the start or end time of an appointment that the user is composing.</span></span> <span data-ttu-id="a1b68-105">これらの非同期メソッドは、アドインの作成のみに使用できます。これらのメソッドを使用するには、「新規[作成フォーム用の outlook アドインを作成](compose-scenario.md)する」で説明されているように、outlook が新規作成フォームでアドインをアクティブにするために、アドインマニフェストが適切にセットアップされていることを確認してください。</span><span class="sxs-lookup"><span data-stu-id="a1b68-105">These asynchronous methods are available to only compose add-ins. To use these methods, make sure you have set up the add-in manifest appropriately for Outlook to activate the add-in in compose forms, as described in [Create Outlook add-ins for compose forms](compose-scenario.md).</span></span>

<span data-ttu-id="a1b68-p102">[start](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) プロパティおよび [end](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) プロパティは、新規作成フォームと閲覧フォームの両方の予定で利用できます。閲覧フォームでは親オブジェクトから直接プロパティにアクセスでき、それには次を使用します。</span><span class="sxs-lookup"><span data-stu-id="a1b68-p102">The [start](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) and [end](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) properties are available for appointments in both compose and read forms. In a read form, you can access the properties directly from the parent object, as in:</span></span>

```js
item.start
```

<span data-ttu-id="a1b68-108">、</span><span class="sxs-lookup"><span data-stu-id="a1b68-108">and in:</span></span>

```js
item.end
```

<span data-ttu-id="a1b68-109">ただし新規作成フォームでは、ユーザーとアドインの両方が同時に時刻を挿入または変更できるため、次に示すように、非同期メソッド **getAsync** を使用して開始時刻または終了時刻を取得する必要があります。</span><span class="sxs-lookup"><span data-stu-id="a1b68-109">But in a compose form, because both the user and your add-in can be inserting or changing the time at the same time, you must use the asynchronous method **getAsync** to get the start or end time, as shown below:</span></span>

```js
item.start.getAsync
```

<span data-ttu-id="a1b68-110">、および</span><span class="sxs-lookup"><span data-stu-id="a1b68-110">and:</span></span>

```js
item.end.getAsync
```

<span data-ttu-id="a1b68-111">Office JavaScript API のほとんどの非同期メソッドと同様に、 **getAsync**および**setasync**はオプションの入力パラメーターを受け取ります。</span><span class="sxs-lookup"><span data-stu-id="a1b68-111">As with most asynchronous methods in the Office JavaScript API, **getAsync** and **setAsync** take optional input parameters.</span></span> <span data-ttu-id="a1b68-112">これらのオプションの入力パラメーターの指定について詳しくは、「 [Office アドインにおける非同期プログラミング](../develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-inline)」の「 [オプションのパラメーターを非同期メソッドに渡す](../develop/asynchronous-programming-in-office-add-ins.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="a1b68-112">For more information about specifying these optional input parameters, see [passing optional parameters to asynchronous methods](../develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-inline) in [Asynchronous programming in Office Add-ins](../develop/asynchronous-programming-in-office-add-ins.md).</span></span>


## <a name="get-the-start-or-end-time"></a><span data-ttu-id="a1b68-113">開始時刻または終了時刻を取得する</span><span class="sxs-lookup"><span data-stu-id="a1b68-113">Get the start or end time</span></span>

<span data-ttu-id="a1b68-p104">このセクションでは、ユーザーが作成している予定の開始時刻を取得して、その時刻を表示するサンプル コードについて説明します。同じコードを使用して、**start** プロパティを **end** プロパティに置き換えると終了時刻を取得できます。このサンプル コードは、以下に示すアドイン マニフェストのルールによって予定の新規作成フォームでアドインがアクティブになることを想定しています。</span><span class="sxs-lookup"><span data-stu-id="a1b68-p104">This section shows a code sample that gets the start time of the appointment that the user is composing and displays the time. You can use the same code and replace the **start** property by the **end** property to get the end time. This code sample assumes a rule in the add-in manifest that activates the add-in in a compose form for an appointment, as shown below.</span></span>


```XML
<Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>

```

<span data-ttu-id="a1b68-p105">**item.start.getAsync** または **item.end.getAsync** を使用する場合は、非同期呼び出しの状態と結果を確認するコールバック メソッドを用意します。_asyncContext_ オプション パラメーターを使用して、コールバック メソッドに必要な引数を指定できます。コールバックの出力パラメーター _asyncResult_ を使用して、状態、結果およびエラーを取得できます。非同期呼び出しに成功すると、**AsyncResult.value** プロパティを使用して開始時刻を UTC 形式の [Date](/javascript/api/office/office.asyncresult#value) オブジェクトとして取得できます。</span><span class="sxs-lookup"><span data-stu-id="a1b68-p105">To use **item.start.getAsync** or **item.end.getAsync**, provide a callback method that checks for the status and result of the asynchronous call. You can provide any necessary arguments to the callback method through the  _asyncContext_ optional parameter. You can obtain status, results and any error using the output parameter _asyncResult_ of the callback. If the asynchronous call is successful, you can get the start time as a **Date** object in UTC format using the [AsyncResult.value](/javascript/api/office/office.asyncresult#value) property.</span></span>


```js
var item;

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


## <a name="set-the-start-or-end-time"></a><span data-ttu-id="a1b68-121">開始時刻または終了時刻を設定する</span><span class="sxs-lookup"><span data-stu-id="a1b68-121">Set the start or end time</span></span>

<span data-ttu-id="a1b68-p106">ここでは、ユーザーが作成している予定またはメッセージの開始時刻を設定するサンプル コードについて説明します。同じコードを使用して、**start** プロパティを **end** プロパティに置き換えると終了時刻を設定できます。予定の新規作成フォームに既存の開始時刻がある場合、後で開始時刻を設定すると、以前の予定の期間が保たれるように終了時刻を調整します。予定の新規作成フォームに既存の終了時刻がある場合、後で終了時刻を設定すると、期間と終了時刻の両方が調整されます。予定が終日イベントとして設定されている場合、開始時刻を設定すると、終了時刻を 24 時間後に調整し、新規作成フォームの終日イベントの UI をオフにします。</span><span class="sxs-lookup"><span data-stu-id="a1b68-p106">This section shows a code sample that sets the start time of the appointment or message that the user is composing. You can use the same code and replace the **start** property by the **end** property to set the end time. Note that if the appointment compose form already has an existing start time, setting the start time subsequently will adjust the end time to maintain any previous duration for the appointment. If the appointment compose form already has an existing end time, setting the end time subsequently will adjust both the duration and end time. If the appointment has been set as an all-day event, setting the start time will adjust the end time to 24 hours later, and uncheck the UI for the all-day event in the compose form.</span></span>

<span data-ttu-id="a1b68-127">前の例と同様、このコード サンプルは、予定の新規作成フォームでアドインをアクティブ化するアドインのマニフェストのルールを想定しています。</span><span class="sxs-lookup"><span data-stu-id="a1b68-127">Similar to the previous example, this code sample assumes a rule in the add-in manifest that activates the add-in in a compose form for an appointment.</span></span>

<span data-ttu-id="a1b68-p107">**item.start.setAsync** または **item.end.setAsync** を使用する場合は、**dateTime** パラメーターに UTC で _Date_ の値を指定します。クライアントでユーザーによる入力に基づいて日付を取得する場合は、[mailbox.convertToUtcClientTime](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) を使用して、値を UTC の **Date** オブジェクトに変換します。オプションのコールバック メソッドと、_asyncContext_ パラメーターでそのコールバック メソッドの引数を指定できます。コールバックの _asyncResult_ 出力パラメーターで、状態、結果およびエラー メッセージを確認する必要があります。非同期呼び出しが成功すると、指定した開始時刻または終了時刻文字列が **setAsync** によってプレーン テキストとして挿入され、そのアイテムの既存の開始時刻または終了時刻が上書きされます。</span><span class="sxs-lookup"><span data-stu-id="a1b68-p107">To use **item.start.setAsync** or **item.end.setAsync**, specify a **Date** value in UTC in the _dateTime_ parameter. If you get a date based on an input by the user on the client, you can use [mailbox.convertToUtcClientTime](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) to convert the value to a **Date** object in UTC. You can provide an optional callback method and any arguments for the callback method in the _asyncContext_ parameter. You should check the status, result and any error message in the _asyncResult_ output parameter of the callback. If the asynchronous call is successful, **setAsync** inserts the specified start or end time string as plain text, overwriting any existing start or end time for that item.</span></span>




```js
var item;

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
    var startDate = new Date("September 27, 2012 12:30:00");
    
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


## <a name="see-also"></a><span data-ttu-id="a1b68-133">関連項目</span><span class="sxs-lookup"><span data-stu-id="a1b68-133">See also</span></span>

- [<span data-ttu-id="a1b68-134">Outlook で新規作成フォームのアイテム データを取得および設定する</span><span class="sxs-lookup"><span data-stu-id="a1b68-134">Get and set item data in a compose form in Outlook</span></span>](get-and-set-item-data-in-a-compose-form.md)    
- [<span data-ttu-id="a1b68-135">閲覧または新規作成フォームの Outlook アイテム データを取得および設定する</span><span class="sxs-lookup"><span data-stu-id="a1b68-135">Get and set Outlook item data in read or compose forms</span></span>](item-data.md)   
- [<span data-ttu-id="a1b68-136">新規作成フォーム用の Outlook アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="a1b68-136">Create Outlook add-ins for compose forms</span></span>](compose-scenario.md)    
- [<span data-ttu-id="a1b68-137">Office アドインにおける非同期プログラミング</span><span class="sxs-lookup"><span data-stu-id="a1b68-137">Asynchronous programming in Office Add-ins</span></span>](../develop/asynchronous-programming-in-office-add-ins.md)
- [<span data-ttu-id="a1b68-138">Outlook の予定またはメッセージを作成するときに受信者を取得、設定、追加する</span><span class="sxs-lookup"><span data-stu-id="a1b68-138">Get, set, or add recipients when composing an appointment or message in Outlook</span></span>](get-set-or-add-recipients.md)  
- [<span data-ttu-id="a1b68-139">Outlook で予定またはメッセージを作成するときに件名を取得または設定する</span><span class="sxs-lookup"><span data-stu-id="a1b68-139">Get or set the subject when composing an appointment or message in Outlook</span></span>](get-or-set-the-subject.md)   
- [<span data-ttu-id="a1b68-140">Outlook で予定またはメッセージを作成するときに本文にデータを挿入する</span><span class="sxs-lookup"><span data-stu-id="a1b68-140">Insert data in the body when composing an appointment or message in Outlook</span></span>](insert-data-in-the-body.md)   
- [<span data-ttu-id="a1b68-141">Outlook で予定を作成するときに場所を取得または設定する</span><span class="sxs-lookup"><span data-stu-id="a1b68-141">Get or set the location when composing an appointment in Outlook</span></span>](get-or-set-the-location-of-an-appointment.md)
    
