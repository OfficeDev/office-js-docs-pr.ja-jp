---
title: Outlook アドインで件名を取得または設定する
description: Outlook アドインで、メッセージまたは予定の件名を取得または設定する方法について説明します。
ms.date: 04/15/2019
localization_priority: Normal
ms.openlocfilehash: b27f6011b1754fa68a1af87f57034e95fd0d54e0
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/20/2020
ms.locfileid: "42166494"
---
# <a name="get-or-set-the-subject-when-composing-an-appointment-or-message-in-outlook"></a><span data-ttu-id="7361d-103">Outlook で予定またはメッセージを作成するときに件名を取得または設定する</span><span class="sxs-lookup"><span data-stu-id="7361d-103">Get or set the subject when composing an appointment or message in Outlook</span></span>

<span data-ttu-id="7361d-p101">JavaScript API for Office には、ユーザーが作成する予定やメッセージの件名を取得および設定する非同期メソッド ([subject.getAsync](/javascript/api/outlook/office.Subject#getasync-options--callback-) と [subject.setAsync](/javascript/api/outlook/office.Subject#setasync-subject--options--callback-)) が用意されています。これらのメソッドを使用する場合は、新規作成フォームでアドインをアクティブ化するようにアドイン マニフェストが Outlook 用に適切にセット アップされていることを確認してください。</span><span class="sxs-lookup"><span data-stu-id="7361d-p101">The JavaScript API for Office provides asynchronous methods ([subject.getAsync](/javascript/api/outlook/office.Subject#getasync-options--callback-) and [subject.setAsync](/javascript/api/outlook/office.Subject#setasync-subject--options--callback-)) to get and set the subject of an appointment or message that the user is composing. These asynchronous methods are available only to compose add-ins. To use these methods, make sure you have set up the add-in manifest appropriately for Outlook to activate the add-in in compose forms.</span></span>

<span data-ttu-id="7361d-p102">**subject** プロパティは、予定とメッセージの新規作成フォームと閲覧フォームの両方で読み取りアクセスで利用できます。閲覧フォームでは、次の例に示すとおり、このプロパティに親オブジェクトから直接アクセスできます。</span><span class="sxs-lookup"><span data-stu-id="7361d-p102">The **subject** property is available for read access in both compose and read forms of appointments and messages. In a read form, you can access the property directly from the parent object, as in:</span></span>

```js
item.subject
```

<span data-ttu-id="7361d-108">ただし、新規作成フォームでは、ユーザーとアドインの両方が同時に件名を挿入または変更できるため、次に示すように、非同期メソッド **getAsync** を使用して件名を取得する必要があります。</span><span class="sxs-lookup"><span data-stu-id="7361d-108">But in a compose form, because both the user and your add-in can be inserting or changing the subject at the same time, you must use the asynchronous method **getAsync** to get the subject, as shown below:</span></span>

```js
item.subject.getAsync
```

<span data-ttu-id="7361d-109">書き込みアクセスでは、**subject** プロパティは新規作成フォームのみで利用でき、閲覧フォームでは利用できません。</span><span class="sxs-lookup"><span data-stu-id="7361d-109">The **subject** property is available for write access in only compose forms and not in read forms.</span></span>

<span data-ttu-id="7361d-p103">JavaScript API for Office のほとんどの非同期メソッドと同じように、**getAsync** と **setAsync** はオプションの入力パラメーターを受け取ります。オプションの入力パラメーターを指定する方法の詳細については、「[Office アドインにおける非同期プログラミング](../develop/asynchronous-programming-in-office-add-ins.md)」の「オプションのパラメーターを非同期メソッドに渡す」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="7361d-p103">As with most asynchronous methods in the JavaScript API for Office, **getAsync** and **setAsync** take optional input parameters. For more information about specifying these optional input parameters, see "Passing optional parameters to asynchronous methods" in [Asynchronous programming in Office Add-ins](../develop/asynchronous-programming-in-office-add-ins.md).</span></span>


## <a name="get-the-subject"></a><span data-ttu-id="7361d-112">件名を取得する</span><span class="sxs-lookup"><span data-stu-id="7361d-112">Get the subject</span></span>

<span data-ttu-id="7361d-p104">このセクションでは、ユーザーが作成している予定またはメッセージの件名を取得して、その件名を表示するサンプル コードについて説明します。このサンプル コードは、以下に示すように、アドイン マニフェストのルールが、予定またはメッセージの新規作成フォームでアドインをアクティブにすることを想定しています。</span><span class="sxs-lookup"><span data-stu-id="7361d-p104">This section shows a code sample that gets the subject of the appointment or message that the user is composing, and displays the subject. This code sample assumes a rule in the add-in manifest that activates the add-in in a compose form for an appointment or message, as shown below.</span></span>


```XML
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit"/>
</Rule>

```

<span data-ttu-id="7361d-p105">**item.subject.getAsync** を使用する場合は、非同期呼び出しの状態と結果を確認するコールバック メソッドを用意します。_asyncContext_ オプション パラメーターを使用して、コールバック メソッドに必要な引数を指定できます。コールバックの出力パラメーター _asyncResult_ を使用して、状態、結果およびエラーを取得できます。非同期呼び出しに成功すると、[AsyncResult.value](/javascript/api/office/office.asyncresult#value) プロパティを使用して件名をプレーン テキスト文字列として取得できます。</span><span class="sxs-lookup"><span data-stu-id="7361d-p105">To use **item.subject.getAsync**, provide a callback method that checks for the status and result of the asynchronous call. You can provide any necessary arguments to the callback method through the  _asyncContext_ optional parameter. You can obtain status, results and any error using the output parameter _asyncResult_ of the callback. If the asynchronous call is successful, you can get the subject as a plain text string using the [AsyncResult.value](/javascript/api/office/office.asyncresult#value) property.</span></span>


```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get the subject of the item being composed.
        getSubject();
    });
}

// Get the subject of the item that the user is composing.
function getSubject() {
    item.subject.getAsync(
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully got the subject, display it.
                write ('The subject is: ' + asyncResult.value);
            }
        });
}

// Write to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## <a name="set-the-subject"></a><span data-ttu-id="7361d-119">件名を設定する</span><span class="sxs-lookup"><span data-stu-id="7361d-119">Set the subject</span></span>


<span data-ttu-id="7361d-p106">このセクションでは、ユーザーが作成している予定またはメッセージの件名を設定するサンプル コードについて説明します。前のサンプルと同様に、このサンプル コードは、アドイン マニフェストのルールが、予定またはメッセージの新規作成フォームでアドインをアクティブにすることを想定しています。</span><span class="sxs-lookup"><span data-stu-id="7361d-p106">This section shows a code sample that sets the subject of the appointment or message that the user is composing. Similar to the previous example, this code sample assumes a rule in the add-in manifest that activates the add-in in a compose form for an appointment or message.</span></span>

<span data-ttu-id="7361d-p107">**item.subject.setAsync** を使用する場合は、データ パラメーターで最大 255 文字の文字列を指定します。オプションで、コールバック メソッドおよび _asyncContext_ パラメーターにそのコールバック メソッドの引数を指定できます。コールバックの _asyncResult_ 出力パラメーターで、状態、結果およびエラー メッセージを確認する必要があります。非同期呼び出しが成功すると、**setAsync** はそのアイテムの既存の件名を上書きして、指定された件名の文字列をプレーン テキストとして挿入します。</span><span class="sxs-lookup"><span data-stu-id="7361d-p107">To use **item.subject.setAsync**, specify a string of up to 255 characters in the data parameter. Optionally, you can provide a callback method and any arguments for the callback method in the  _asyncContext_ parameter. You should check the status, result and any error message in the _asyncResult_ output parameter of the callback. If the asynchronous call is successful, **setAsync** inserts the specified subject string as plain text, overwriting any existing subject for that item.</span></span>

```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set the subject of the item being composed.
        setSubject();
    });
}

// Set the subject of the item that the user is composing.
function setSubject() {
    var today = new Date();
    var subject;

    // Customize the subject with today's date.
    subject = 'Summary for ' + today.toLocaleDateString();

    item.subject.setAsync(
        subject,
        { asyncContext: { var1: 1, var2: 2 } },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully set the subject.
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


## <a name="see-also"></a><span data-ttu-id="7361d-126">関連項目</span><span class="sxs-lookup"><span data-stu-id="7361d-126">See also</span></span>

- [<span data-ttu-id="7361d-127">Outlook で新規作成フォームのアイテム データを取得および設定する</span><span class="sxs-lookup"><span data-stu-id="7361d-127">Get and set item data in a compose form in Outlook</span></span>](get-and-set-item-data-in-a-compose-form.md)   
- [<span data-ttu-id="7361d-128">閲覧または新規作成フォームの Outlook アイテム データを取得および設定する</span><span class="sxs-lookup"><span data-stu-id="7361d-128">Get and set Outlook item data in read or compose forms</span></span>](item-data.md)    
- [<span data-ttu-id="7361d-129">新規作成フォーム用の Outlook アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="7361d-129">Create Outlook add-ins for compose forms</span></span>](compose-scenario.md)    
- [<span data-ttu-id="7361d-130">Office アドインにおける非同期プログラミング</span><span class="sxs-lookup"><span data-stu-id="7361d-130">Asynchronous programming in Office Add-ins</span></span>](../develop/asynchronous-programming-in-office-add-ins.md)
- [<span data-ttu-id="7361d-131">Outlook の予定またはメッセージを作成するときに受信者を取得、設定、追加する</span><span class="sxs-lookup"><span data-stu-id="7361d-131">Get, set, or add recipients when composing an appointment or message in Outlook</span></span>](get-set-or-add-recipients.md)  
- [<span data-ttu-id="7361d-132">Outlook で予定またはメッセージを作成するときに本文にデータを挿入する</span><span class="sxs-lookup"><span data-stu-id="7361d-132">Insert data in the body when composing an appointment or message in Outlook</span></span>](insert-data-in-the-body.md)   
- [<span data-ttu-id="7361d-133">Outlook で予定を作成するときに場所を取得または設定する</span><span class="sxs-lookup"><span data-stu-id="7361d-133">Get or set the location when composing an appointment in Outlook</span></span>](get-or-set-the-location-of-an-appointment.md) 
- [<span data-ttu-id="7361d-134">Outlook で予定を作成するときに時刻を取得または設定する</span><span class="sxs-lookup"><span data-stu-id="7361d-134">Get or set the time when composing an appointment in Outlook</span></span>](get-or-set-the-time-of-an-appointment.md)
    
