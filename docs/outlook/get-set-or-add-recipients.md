---
title: Outlook アドインで受信者を取得または変更する
description: Outlook アドインで、メッセージまたは予定の受信者を取得、設定、追加する方法について説明します。
ms.date: 06/27/2022
ms.localizationpriority: medium
ms.openlocfilehash: 0cd51bd20d90d0183cbe0794b644eaa307e2d6b7
ms.sourcegitcommit: 2a0bd3155c732b6010b7e1e612f4bd7e45f79c52
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/29/2022
ms.locfileid: "66241289"
---
# <a name="get-set-or-add-recipients-when-composing-an-appointment-or-message-in-outlook"></a>Outlook の予定またはメッセージを作成するときに受信者を取得、設定、追加する

Office JavaScript API は、非同期メソッド ([Recipients.getAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-getasync-member(1))、 [Recipients.setAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-setasync-member(1))、 [Recipients.addAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-addasync-member(1))) を提供し、それぞれ、予定またはメッセージの作成形式で受信者を取得、設定、または追加します。 これらの非同期メソッドは、アドインを作成する場合にのみ使用できます。これらのメソッドを使用するには、「作成フォーム用の Outlook アドインを作成する」の説明に従って、Outlook がアドイン作成フォームをアクティブ化するために適切 [にアドイン](compose-scenario.md) マニフェストを設定していることを確認します。

予定やメッセージ内の受信者を表すプロパティの一部は、新規作成フォームと閲覧フォームで読み取りアクセスで使用できます。この種のプロパティには、予定の [optionalAttendees](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) と [requiredAttendees](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)、メッセージの [cc](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) と [to](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) が含まれます。

閲覧フォームでは、次に示すように親オブジェクトから直接プロパティにアクセスできます:

```js
item.cc
```

ただし、作成フォームでは、ユーザーとアドインの両方が同時に受信者を挿入または変更できるため、次の例のように、非同期メソッド `getAsync` を使用してこれらのプロパティを取得する必要があります。

```js
item.cc.getAsync
```

これらのプロパティを書き込みアクセスに使用できるのは新規作成フォームに限られ、閲覧フォームでは使用できません。

JavaScript API for Office のほとんどの非同期メソッドと同様に、`getAsync``setAsync`省略可能な入力パラメーターを`addAsync`受け取ります。 これらのオプションの入力パラメーターの指定について詳しくは、「 [Office アドインにおける非同期プログラミング](../develop/asynchronous-programming-in-office-add-ins.md#pass-optional-parameters-inline)」の「 [オプションのパラメーターを非同期メソッドに渡す](../develop/asynchronous-programming-in-office-add-ins.md)」を参照してください。

## <a name="get-recipients"></a>受信者を取得する

このセクションでは、新規作成する予定やメッセージの受信者を取得し、その受信者の電子メール アドレスを表示するコード例を示します。下記のように、このコード例は、予定やメッセージ用の新規作成フォームでアドインをアクティブ化するルールがアドイン マニフェスト内にあることを前提にしています。

```XML
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit"/>
</Rule>
```

Office JavaScript API では、予定の受信者を表すプロパティ ( **optionalAttendees** と **requiredAttendees**) はメッセージのプロパティ ([bcc](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)、 **cc**、to) とは異なるため、最初 **に** [item.itemType](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) プロパティを使用して、構成されるアイテムが予定またはメッセージであるかどうかを識別する必要があります。 作成モードでは、予定とメッセージのこれらのプロパティはすべて [Recipients](/javascript/api/outlook/office.recipients) オブジェクトであるため、非同期メソッドを適用して、 `Recipients.getAsync`対応する受信者を取得できます。

使用 `getAsync`するには、状態、結果、および非同期 `getAsync` 呼び出しによって返されるエラーを確認するコールバック メソッドを指定します。 オプションの _asyncContext_ パラメーターを使用して、コールバック メソッドに引数を指定できます。 The callback method returns an _asyncResult_ output parameter. [AsyncResult](/javascript/api/office/office.asyncresult) パラメーター オブジェクトのプロパティと`error`プロパティを使用`status`して、非同期呼び出しの状態とエラー メッセージ、および実際の受信者を`value`取得するプロパティを確認できます。 受信者は、[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) オブジェクトの配列として表されます。

メソッドは非同期であるため `getAsync` 、受信者を正常に取得することに依存する後続のアクションがある場合は、非同期呼び出しが正常に完了したときに、対応するコールバック メソッドでのみこのようなアクションを開始するようにコードを整理する必要があることに注意してください。

> [!IMPORTANT]
> このメソッドは `getAsync` 、Outlook クライアントによって解決された受信者のみを返します。 解決された受信者には、次の特性があります。
>
> - 受信者が送信者のアドレス帳に保存されたエントリを持っている場合、Outlook は電子メール アドレスを受信者の保存された表示名に解決します。
> - Teams 会議の状態アイコンは、受信者の名前または電子メール アドレスの前に表示されます。
> - 受信者の名前または電子メール アドレスの後にセミコロンが表示されます。
> - 受信者の名前または電子メール アドレスは、下線が付けられているか、ボックスに囲まれています。
>
> メール アイテムに追加されたメール アドレスを解決するには、送信者が **Tab** キーを使用するか、オートコンプリート リストから推奨される連絡先または電子メール アドレスを選択する必要があります。

> [!NOTE]
> Outlook on the webおよび Windows では、ユーザーが連絡先またはプロファイル カードから連絡先のメール アドレス リンクをアクティブ化して新しいメッセージを作成した場合、アドインの`Recipients.getAsync`呼び出しは、連絡先の保存された名前ではなく、関連付けられた`EmailAddressDetails`オブジェクトのプロパティで`displayName`連絡先の電子メール アドレスを返します。
>
> 詳細については、 [関連する GitHub の問題](https://github.com/OfficeDev/office-js/issues/2201)を参照してください。

```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get all the recipients of the composed item.
        getAllRecipients();
    });
}

// Get the email addresses of all the recipients of the composed item.
function getAllRecipients() {
    // Local objects to point to recipients of either
    // the appointment or message that is being composed.
    // bccRecipients applies to only messages, not appointments.
    var toRecipients, ccRecipients, bccRecipients;
    // Verify if the composed item is an appointment or message.
    if (item.itemType == Office.MailboxEnums.ItemType.Appointment) {
        toRecipients = item.requiredAttendees;
        ccRecipients = item.optionalAttendees;
    }
    else {
        toRecipients = item.to;
        ccRecipients = item.cc;
        bccRecipients = item.bcc;
    }
    
    // Use asynchronous method getAsync to get each type of recipients
    // of the composed item. Each time, this example passes an anonymous 
    // callback function that doesn't take any parameters.
    toRecipients.getAsync(function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed){
            write(asyncResult.error.message);
        }
        else {
            // Async call to get to-recipients of the item completed.
            // Display the email addresses of the to-recipients. 
            write ('To-recipients of the item:');
            displayAddresses(asyncResult);
        }    
    }); // End getAsync for to-recipients.

    // Get any cc-recipients.
    ccRecipients.getAsync(function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed){
            write(asyncResult.error.message);
        }
        else {
            // Async call to get cc-recipients of the item completed.
            // Display the email addresses of the cc-recipients.
            write ('Cc-recipients of the item:');
            displayAddresses(asyncResult);
        }
    }); // End getAsync for cc-recipients.

    // If the item has the bcc field, i.e., item is message,
    // get any bcc-recipients.
    if (bccRecipients) {
        bccRecipients.getAsync(function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed){
            write(asyncResult.error.message);
        }
        else {
            // Async call to get bcc-recipients of the item completed.
            // Display the email addresses of the bcc-recipients.
            write ('Bcc-recipients of the item:');
            displayAddresses(asyncResult);
        }
                        
        }); // End getAsync for bcc-recipients.
     }
}

// Recipients are in an array of EmailAddressDetails
// objects passed in asyncResult.value.
function displayAddresses (asyncResult) {
    for (var i=0; i<asyncResult.value.length; i++)
        write (asyncResult.value[i].emailAddress);
}

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

## <a name="set-recipients"></a>受信者を設定する

このセクションでは、ユーザーが新規作成する予定やメッセージの受信者を設定するコード例を示しています。 受信者を設定すると、既存の受信者が上書きされます。 この例では、前述の新規作成フォームで受信者を取得する例と同様に、アドインが予定とメッセージの新規作成フォームでアクティブ化されることを想定しています。 この例では、まず、構成されたアイテムが予定またはメッセージであるかどうかを確認し、非同期メソッドを適用して、 `Recipients.setAsync`予定またはメッセージの受信者を表す適切なプロパティに適用します。

呼び出す `setAsync`場合は、次のいずれかの形式で  _、recipients_ パラメーターの入力引数として配列を指定します。

- SMTP アドレスである文字列の配列。
- 辞書の配列。次のコード例に示されているように、それぞれ表示名と電子メール アドレスが含まれています。
- メソッドによって返されるオブジェクトと同様のオブジェクトの`getAsync`配列`EmailAddressDetails`。
  
必要に応じて、コールバック メソッドをメソッドの `setAsync` 入力引数として指定して、受信者を正常に設定することに依存するコードが、それが発生した場合にのみ実行されるようにすることができます。 オプションの _asyncContext_ パラメーターを使用してコールバック メソッドの引数を提供することもできます。 コールバック メソッドを使用する場合は、 _asyncResult_ 出力パラメーターにアクセスし、パラメーター オブジェクトの **状態** プロパティと **エラー** プロパティを `AsyncResult` 使用して、非同期呼び出しの状態とエラー メッセージを確認できます。

```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set recipients of the composed item.
        setRecipients();
    });
}

// Set the display name and email addresses of the recipients of 
// the composed item.
function setRecipients() {
    // Local objects to point to recipients of either
    // the appointment or message that is being composed.
    // bccRecipients applies to only messages, not appointments.
    var toRecipients, ccRecipients, bccRecipients;

    // Verify if the composed item is an appointment or message.
    if (item.itemType == Office.MailboxEnums.ItemType.Appointment) {
        toRecipients = item.requiredAttendees;
        ccRecipients = item.optionalAttendees;
    }
    else {
        toRecipients = item.to;
        ccRecipients = item.cc;
        bccRecipients = item.bcc;
    }
    
    // Use asynchronous method setAsync to set each type of recipients
    // of the composed item. Each time, this example passes a set of
    // names and email addresses to set, and an anonymous 
    // callback function that doesn't take any parameters. 
    toRecipients.setAsync(
        [{
            "displayName":"Graham Durkin", 
            "emailAddress":"graham@contoso.com"
         },
         {
            "displayName" : "Donnie Weinberg",
            "emailAddress" : "donnie@contoso.com"
         }],
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Async call to set to-recipients of the item completed.

            }    
    }); // End to setAsync.


    // Set any cc-recipients.
    ccRecipients.setAsync(
        [{
             "displayName":"Perry Horning", 
             "emailAddress":"perry@contoso.com"
         },
         {
             "displayName" : "Guy Montenegro",
             "emailAddress" : "guy@contoso.com"
         }],
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Async call to set cc-recipients of the item completed.
            }
    }); // End cc setAsync.


    // If the item has the bcc field, i.e., item is message,
    // set bcc-recipients.
    if (bccRecipients) {
        bccRecipients.setAsync(
            [{
                 "displayName":"Lewis Cate", 
                 "emailAddress":"lewis@contoso.com"
             },
             {
                 "displayName" : "Francisco Stitt",
                 "emailAddress" : "francisco@contoso.com"
             }],
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Failed){
                    write(asyncResult.error.message);
                }
                else {
                    // Async call to set bcc-recipients of the item completed.
                    // Do whatever appropriate for your scenario.
                }
        }); // End bcc setAsync.
    }
}

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}

```

## <a name="add-recipients"></a>受信者を追加する

予定またはメッセージ内の既存の受信者を上書きしない場合は、非同期メソッドを使用 `Recipients.setAsync`して受信者を `Recipients.addAsync` 追加できます。 `addAsync`は、受信者の入力引数が必要であるという点と同様`setAsync`_に機能します_。 オプションで、コールバック メソッドを指定し、asyncContext パラメーターを使用してコールバックの引数を提供できます。 その後、コールバック メソッドの _asyncResult_ 出力パラメーターを使用して、非同期`addAsync`呼び出しの状態、結果、エラーを確認できます。 次の例は、新規作成されるアイテムが予定かどうかチェックし、その予定に 2 人の必須の出席者を付加します。

```js
// Add specified recipients as required attendees of
// the composed appointment. 
function addAttendees() {
    if (item.itemType == Office.MailboxEnums.ItemType.Appointment) {
        item.requiredAttendees.addAsync(
        [{
            "displayName":"Kristie Jensen", 
            "emailAddress":"kristie@contoso.com"
         },
         {
            "displayName" : "Pansy Valenzuela",
            "emailAddress" : "pansy@contoso.com"
          }],
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Async call to add attendees completed.
                // Do whatever appropriate for your scenario.
            }
        }); // End addAsync.
    }
}
```

## <a name="see-also"></a>関連項目

- [Outlook で新規作成フォームのアイテム データを取得および設定する](get-and-set-item-data-in-a-compose-form.md)
- [閲覧または新規作成フォームの Outlook アイテム データを取得および設定する](item-data.md)
- [新規作成フォーム用の Outlook アドインを作成する](compose-scenario.md)
- [Office アドインにおける非同期プログラミング](../develop/asynchronous-programming-in-office-add-ins.md)
- [Outlook で予定またはメッセージを作成するときに件名を取得または設定する](get-or-set-the-subject.md)
- [Outlook で予定またはメッセージを作成するときに本文にデータを挿入する](insert-data-in-the-body.md)
- [Outlook で予定を作成するときに場所を取得または設定する](get-or-set-the-location-of-an-appointment.md)
- [Outlook で予定を作成するときに時刻を取得または設定する](get-or-set-the-time-of-an-appointment.md)
