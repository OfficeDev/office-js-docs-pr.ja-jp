---
title: Outlook アドインで件名を取得または設定する
description: Outlook アドインで、メッセージまたは予定の件名を取得または設定する方法について説明します。
ms.date: 10/07/2022
ms.localizationpriority: medium
ms.openlocfilehash: 79e38a310bf62eae55ef020c2f6c978ace824255
ms.sourcegitcommit: a2df9538b3deb32ae3060ecb09da15f5a3d6cb8d
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/12/2022
ms.locfileid: "68541130"
---
# <a name="get-or-set-the-subject-when-composing-an-appointment-or-message-in-outlook"></a>Outlook で予定またはメッセージを作成するときに件名を取得または設定する

Office JavaScript API は、ユーザーが作成している予定またはメッセージの件名を取得および設定するための非同期メソッド ([subject.getAsync](/javascript/api/outlook/office.subject#outlook-office-subject-getasync-member(1)) と [subject.setAsync](/javascript/api/outlook/office.subject#outlook-office-subject-setasync-member(1))) を提供します。 これらの非同期メソッドは、アドインを作成する場合にのみ使用できます。これらのメソッドを使用するには、アドイン [の作成フォームをアクティブ化](compose-scenario.md)するために、Outlook 用にアドイン XML マニフェストを適切に設定していることを確認します。 Office アドイン [の Teams マニフェスト (プレビュー) を](../develop/json-manifest-overview.md)使用するアドインでは、アクティブ化ルールはサポートされていません。

The **subject** property is available for read access in both compose and read forms of appointments and messages. In a read form, you can access the property directly from the parent object, as in:

```js
item.subject
```

ただし、新規作成フォームでは、ユーザーとアドインの両方が同時に件名を挿入または変更できるため、次に示すように、非同期メソッド **getAsync** を使用して件名を取得する必要があります。

```js
item.subject.getAsync
```

書き込みアクセスでは、**subject** プロパティは新規作成フォームのみで利用でき、閲覧フォームでは利用できません。

Office JavaScript API のほとんどの非同期メソッドと同様に、 **getAsync** と **setAsync** は省略可能な入力パラメーターを受け取ります。 オプションの入力パラメーターを指定する方法の詳細については、「[Office アドインにおける非同期プログラミング](../develop/asynchronous-programming-in-office-add-ins.md)」を参照してください。

## <a name="get-the-subject"></a>件名を取得する

このセクションでは、ユーザーが作成している予定またはメッセージの件名を取得して、その件名を表示するサンプル コードについて説明します。 このサンプル コードは、以下に示すように、アドイン マニフェストのルールが、予定またはメッセージの新規作成フォームでアドインをアクティブにすることを想定しています。 Office アドイン [の Teams マニフェスト (プレビュー)](../develop/json-manifest-overview.md) を使用するアドインでは、アクティブ化規則はサポートされていません。

```XML
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit"/>
</Rule>
```

**item.subject.getAsync** を使用するには、非同期呼び出しの状態と結果を確認するコールバック関数を指定します。 _asyncContext_ 省略可能なパラメーターを使用して、コールバック関数に必要な引数を指定できます。 You can obtain status, results and any error using the output parameter _asyncResult_ of the callback. If the asynchronous call is successful, you can get the subject as a plain text string using the [AsyncResult.value](/javascript/api/office/office.asyncresult#office-office-asyncresult-value-member) property.

```js
let item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready method.
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

## <a name="set-the-subject"></a>件名を設定する

このセクションでは、ユーザーが作成している予定またはメッセージの件名を設定するサンプル コードについて説明します。 前のサンプルと同様に、このサンプル コードは、アドイン マニフェストのルールが、予定またはメッセージの新規作成フォームでアドインをアクティブにすることを想定しています。 Office アドイン [の Teams マニフェスト (プレビュー)](../develop/json-manifest-overview.md) を使用するアドインでは、アクティブ化規則はサポートされていません。

**item.subject.setAsync** を使用するには、データ パラメーターに最大 255 文字の文字列を指定します。 必要に応じて、  _asyncContext_ パラメーターでコールバック関数とコールバック関数の任意の引数を指定できます。 コールバックの _asyncResult_ 出力パラメーターで、状態、結果およびエラー メッセージを確認する必要があります。 非同期呼び出しが成功すると、**setAsync** はそのアイテムの既存の件名を上書きして、指定された件名の文字列をプレーン テキストとして挿入します。

```js
let item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready method.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set the subject of the item being composed.
        setSubject();
    });
}

// Set the subject of the item that the user is composing.
function setSubject() {
    const today = new Date();
    let subject;

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

## <a name="see-also"></a>関連項目

- [Outlook で新規作成フォームのアイテム データを取得および設定する](get-and-set-item-data-in-a-compose-form.md)
- [閲覧または新規作成フォームの Outlook アイテム データを取得および設定する](item-data.md)
- [新規作成フォーム用の Outlook アドインを作成する](compose-scenario.md)
- [Office アドインにおける非同期プログラミング](../develop/asynchronous-programming-in-office-add-ins.md)
- [Outlook の予定またはメッセージを作成するときに受信者を取得、設定、追加する](get-set-or-add-recipients.md)  
- [Outlook で予定またはメッセージを作成するときに本文にデータを挿入する](insert-data-in-the-body.md)
- [Outlook で予定を作成するときに場所を取得または設定する](get-or-set-the-location-of-an-appointment.md)
- [Outlook で予定を作成するときに時刻を取得または設定する](get-or-set-the-time-of-an-appointment.md)
