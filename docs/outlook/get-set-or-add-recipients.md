---
title: Outlook アドインで受信者を取得または変更する
description: Outlook アドインで、メッセージまたは予定の受信者を取得、設定、追加する方法について説明します。
ms.date: 12/10/2019
localization_priority: Normal
ms.openlocfilehash: d6e69b3adc8ddc9f5606e3ec522c56a621eb3664
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609127"
---
# <a name="get-set-or-add-recipients-when-composing-an-appointment-or-message-in-outlook"></a>Outlook の予定またはメッセージを作成するときに受信者を取得、設定、追加する


Office JavaScript API は、非同期メソッド ([getAsync](/javascript/api/outlook/office.Recipients#getasync-options--callback-)、 [recipients async](/javascript/api/outlook/office.Recipients#setasync-recipients--options--callback-)、または[recipients](/javascript/api/outlook/office.Recipients#addasync-recipients--options--callback-)) を使用して、予定またはメッセージの新規作成フォームで受信者を取得、設定、または追加します。 これらの非同期メソッドは、アドインの作成のみに使用できます。これらのメソッドを使用するには、「新規[作成フォーム用の outlook アドインを作成](compose-scenario.md)する」で説明されているように、outlook が新規作成フォームでアドインをアクティブにするために、アドインマニフェストが適切にセットアップされていることを確認してください。

予定やメッセージ内の受信者を表すプロパティの一部は、新規作成フォームと閲覧フォームで読み取りアクセスで使用できます。この種のプロパティには、予定の [optionalAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) と [requiredAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)、メッセージの [cc](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) と [to](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) が含まれます。

閲覧フォームでは、次に示すように親オブジェクトから直接プロパティにアクセスできます:

```js
item.cc
```

しかし、新規作成フォームでは、ユーザーとアドインの両方が同時に受信者を挿入または変更できるため、次の例に示すように、非同期メソッドを使用してこれらのプロパティを取得する必要があり `getAsync` ます。


```js
item.cc.getAsync
```

これらのプロパティを書き込みアクセスに使用できるのは新規作成フォームに限られ、閲覧フォームでは使用できません。

JavaScript API for Office のほとんどの非同期メソッドと同様に、、、 `getAsync` `setAsync` オプションの `addAsync` 入力パラメーターを取ります。 これらのオプションの入力パラメーターの指定について詳しくは、「 [Office アドインにおける非同期プログラミング](../develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-inline)」の「 [オプションのパラメーターを非同期メソッドに渡す](../develop/asynchronous-programming-in-office-add-ins.md)」を参照してください。


## <a name="get-recipients"></a>受信者を取得する


このセクションでは、新規作成する予定やメッセージの受信者を取得し、その受信者の電子メール アドレスを表示するコード例を示します。下記のように、このコード例は、予定やメッセージ用の新規作成フォームでアドインをアクティブ化するルールがアドイン マニフェスト内にあることを前提にしています。


```XML
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit"/>
</Rule>
```

Office JavaScript API では、予定の受信者を表すプロパティ ([必須**出席者**] および **[requiredat]**) は、メッセージ ([bcc](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)、 **cc**、および**to**) とは別のものであるため、最初に[アイテムの itemType](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)プロパティを使用して、作成されているアイテムが予定またはメッセージであるかどうかを識別します。 新規作成モードでは、予定およびメッセージのすべてのプロパティが[Recipients](/javascript/api/outlook/office.Recipients)オブジェクトであるため、非同期メソッドを適用して `Recipients.getAsync` 対応する受信者を取得できます。

を使用し `getAsync` て、非同期呼び出しによって返された状態、結果、およびエラーを確認するコールバックメソッドを提供 `getAsync` します。 オプションの _asyncContext_ パラメーターを使用して、コールバック メソッドに引数を指定できます。 The callback method returns an _asyncResult_ output parameter. `status` `error` [AsyncResult](/javascript/api/office/office.asyncresult) parameter オブジェクトのプロパティとプロパティを使用して、非同期呼び出しの状態およびエラーメッセージを確認し、プロパティを使用して実際の受信者を取得することができ `value` ます。 受信者は、[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) オブジェクトの配列として表されます。

`getAsync`このメソッドは非同期であるため、受信者を正常に取得することに依存する後続のアクションがある場合は、非同期呼び出しが正常に完了したときに、対応するコールバックメソッドでのみそのようなアクションを開始するようにコードを整理する必要があります。




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


このセクションでは、ユーザーが新規作成する予定やメッセージの受信者を設定するコード例を示しています。 受信者を設定すると、既存の受信者が上書きされます。 この例では、前述の新規作成フォームで受信者を取得する例と同様に、アドインが予定とメッセージの新規作成フォームでアクティブ化されることを想定しています。 この例では、最初に、構成されたアイテムが予定またはメッセージかどうかを確認し、非同期メソッドを適用するために、 `Recipients.setAsync` 予定またはメッセージの受信者を表す適切なプロパティを使用します。

を呼び出すときに、 `setAsync` 次のいずれかの形式で、 _recipients_パラメーターの入力引数として配列を指定します。


- SMTP アドレスである文字列の配列。
    
- 辞書の配列。次のコード例に示されているように、それぞれ表示名と電子メール アドレスが含まれています。
    
- オブジェクトの配列 `EmailAddressDetails` 。メソッドによって返される配列に似て `getAsync` います。
    
必要に応じて、メソッドへの入力引数としてコールバックメソッドを指定して、 `setAsync` 受信者の設定に依存するコードが発生した場合にのみ実行するようにすることができます。 オプションの _asyncContext_ パラメーターを使用してコールバック メソッドの引数を提供することもできます。 コールバックメソッドを使用する場合は、 _asyncResult_出力パラメーターにアクセスして、parameter オブジェクトの**status**および**error**プロパティを使用して、 `AsyncResult` 非同期呼び出しの状態およびエラーメッセージを確認できます。




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

予定またはメッセージ内の既存の受信者を上書きしない場合は、を使用する代わりに、 `Recipients.setAsync` `Recipients.addAsync` 非同期メソッドを使用して受信者を追加できます。 `addAsync`は、 `setAsync` _受信者_の入力引数を必要とするのと同じように動作します。 オプションで、コールバック メソッドを指定し、asyncContext パラメーターを使用してコールバックの引数を提供できます。 その後、 `addAsync` コールバックメソッドの_asyncResult_出力パラメーターを使用して、非同期呼び出しの状態、結果、およびエラーを確認できます。 次の例は、新規作成されるアイテムが予定かどうかチェックし、その予定に 2 人の必須の出席者を付加します。


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
    
