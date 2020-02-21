---
title: Outlook アドインで受信者を取得または変更する
description: Outlook アドインで、メッセージまたは予定の受信者を取得、設定、追加する方法について説明します。
ms.date: 12/10/2019
localization_priority: Normal
ms.openlocfilehash: 36849b0ebb7e1dff34d59305d265294452bf395d
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/20/2020
ms.locfileid: "42166576"
---
# <a name="get-set-or-add-recipients-when-composing-an-appointment-or-message-in-outlook"></a><span data-ttu-id="169ab-103">Outlook の予定またはメッセージを作成するときに受信者を取得、設定、追加する</span><span class="sxs-lookup"><span data-stu-id="169ab-103">Get, set, or add recipients when composing an appointment or message in Outlook</span></span>


<span data-ttu-id="169ab-104">JavaScript API for Office には、予定またはメッセージの新規作成フォームで受信者を取得、設定、または追加する非同期メソッド ([getAsync](/javascript/api/outlook/office.Recipients#getasync-options--callback-)、 [recipients](/javascript/api/outlook/office.Recipients#setasync-recipients--options--callback-)、または[受信者. addasync](/javascript/api/outlook/office.Recipients#addasync-recipients--options--callback-)) が用意されています。</span><span class="sxs-lookup"><span data-stu-id="169ab-104">The JavaScript API for Office provides asynchronous methods ([Recipients.getAsync](/javascript/api/outlook/office.Recipients#getasync-options--callback-), [Recipients.setAsync](/javascript/api/outlook/office.Recipients#setasync-recipients--options--callback-), or [Recipients.addAsync](/javascript/api/outlook/office.Recipients#addasync-recipients--options--callback-)) to respectively get, set, or add recipients in a compose form of an appointment or message.</span></span> <span data-ttu-id="169ab-105">これらの非同期メソッドは、アドインの作成のみに使用できます。これらのメソッドを使用するには、「新規[作成フォーム用の outlook アドインを作成](compose-scenario.md)する」で説明されているように、outlook が新規作成フォームでアドインをアクティブにするために、アドインマニフェストが適切にセットアップされていることを確認してください。</span><span class="sxs-lookup"><span data-stu-id="169ab-105">These asynchronous methods are available to only compose add-ins. To use these methods, make sure you have set up the add-in manifest appropriately for Outlook to activate the add-in in compose forms, as described in [Create Outlook add-ins for compose forms](compose-scenario.md).</span></span>

<span data-ttu-id="169ab-p102">予定やメッセージ内の受信者を表すプロパティの一部は、新規作成フォームと閲覧フォームで読み取りアクセスで使用できます。この種のプロパティには、予定の [optionalAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) と [requiredAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)、メッセージの [cc](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) と [to](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) が含まれます。</span><span class="sxs-lookup"><span data-stu-id="169ab-p102">Some of the properties that represent recipients in an appointment or message are available for read access in a compose form and in a read form. These properties include  [optionalAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) and [requiredAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) for appointments, and [cc](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties), and  [to](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) for messages.</span></span>

<span data-ttu-id="169ab-108">閲覧フォームでは、次に示すように親オブジェクトから直接プロパティにアクセスできます:</span><span class="sxs-lookup"><span data-stu-id="169ab-108">In a read form, you can access the property directly from the parent object, such as:</span></span>

```js
item.cc
```

<span data-ttu-id="169ab-109">しかし、新規作成フォームでは、ユーザーとアドインの両方が同時に受信者の挿入や変更を行うことがあるので、以下の例のように非同期メソッド **getAsync** を使用してこれらのプロパティを取得する必要があります。</span><span class="sxs-lookup"><span data-stu-id="169ab-109">But in a compose form, because both the user and your add-in can be inserting or changing a recipient at the same time, you must use the asynchronous method **getAsync** to get these properties, as in the following example:</span></span>


```js
item.cc.getAsync
```

<span data-ttu-id="169ab-110">これらのプロパティを書き込みアクセスに使用できるのは新規作成フォームに限られ、閲覧フォームでは使用できません。</span><span class="sxs-lookup"><span data-stu-id="169ab-110">These properties are available for write access in only compose forms and not read forms.</span></span>

<span data-ttu-id="169ab-p103">JavaScript API for Office のほとんどの非同期メソッドと同じように、**getAsync**、**setAsync**、および **addAsync** はオプションの入力パラメーターを受け取ります。これらのオプション入力パラメーターを指定する方法の詳細については、「[Office アドインにおける非同期プログラミング](../develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-inline)」の「[オプションのパラメーターを非同期メソッドに渡す](../develop/asynchronous-programming-in-office-add-ins.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="169ab-p103">As with most asynchronous methods in the JavaScript API for Office, **getAsync**, **setAsync**, and **addAsync** take optional input parameters. For more information about specifying these optional input parameters, see [passing optional parameters to asynchronous methods](../develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-inline) in [Asynchronous programming in Office Add-ins](../develop/asynchronous-programming-in-office-add-ins.md).</span></span>


## <a name="get-recipients"></a><span data-ttu-id="169ab-113">受信者を取得する</span><span class="sxs-lookup"><span data-stu-id="169ab-113">Get recipients</span></span>


<span data-ttu-id="169ab-p104">このセクションでは、新規作成する予定やメッセージの受信者を取得し、その受信者の電子メール アドレスを表示するコード例を示します。下記のように、このコード例は、予定やメッセージ用の新規作成フォームでアドインをアクティブ化するルールがアドイン マニフェスト内にあることを前提にしています。</span><span class="sxs-lookup"><span data-stu-id="169ab-p104">This section shows a code sample that gets the recipients of the appointment or message that is being composed, and displays the email addresses of the recipients. The code sample assumes a rule in the add-in manifest that activates the add-in in a compose form for an appointment or message, as shown below.</span></span>


```XML
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit"/>
</Rule>
```

<span data-ttu-id="169ab-p105">JavaScript API for Office では、予定の受信者を表すプロパティ (**optionalAttendees** および **requiredAttendees**) とメッセージの受信者を表すプロパティ ([bcc](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)、**cc**、および **to**) は違うので、最初に [item.itemType](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) プロパティを使用して、新規作成するアイテムが予定かメッセージかを識別します。新規作成モードでは、これらの予定とメッセージのプロパティはすべて [Recipients](/javascript/api/outlook/office.Recipients) オブジェクトなので、その後に非同期メソッド **Recipients.getAsync** を適用して対応する受信者を取得できます。</span><span class="sxs-lookup"><span data-stu-id="169ab-p105">In the JavaScript API for Office, because the properties that represent the recipients of an appointment ( **optionalAttendees** and **requiredAttendees**) are different from those of a message ([bcc](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties), **cc**, and **to**), you should first use the [item.itemType](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) property to identify whether the item being composed is an appointment or message. In compose mode, all these properties of appointments and messages are [Recipients](/javascript/api/outlook/office.Recipients) objects, so you can then apply the asynchronous method, **Recipients.getAsync**, to get the corresponding recipients.</span></span>

<span data-ttu-id="169ab-p106">**getAsync** を使用する場合は、非同期 **getAsync** 呼び出しによって返される状態、結果、およびエラーをチェックするコールバック メソッドを用意します。オプションの _asyncContext_ パラメーターを使用して、コールバック メソッドに引数を指定できます。コールバック メソッドは _asyncResult_ 出力パラメーターを返します。**AsyncResult** パラメーター オブジェクトの **status** プロパティと [error](/javascript/api/office/office.asyncresult) プロパティを使用すると、非同期呼び出しの状態とエラー メッセージを確認できます。また、**value** プロパティを使用すると実際の受信者を取得できます。受信者は、[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) オブジェクトの配列として表されます。</span><span class="sxs-lookup"><span data-stu-id="169ab-p106">To use **getAsync**, provide a callback method to check for the status, results, and any error returned by the asynchronous **getAsync** call. You can provide any arguments to the callback method using the optional _asyncContext_ parameter. The callback method returns an _asyncResult_ output parameter. You can use the **status** and **error** properties of the [AsyncResult](/javascript/api/office/office.asyncresult) parameter object to check for status and any error messages of the asynchronous call, and the **value** property to get the actual recipients. Recipients are represented as an array of [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) objects.</span></span>

<span data-ttu-id="169ab-123">**getAsync** メソッドは非同期なので、受信者が正常に取得されることに依存するアクションが後に続く場合、非同期呼び出しが正常に完了した場合にのみ対応するコールバック メソッドでこの種のアクションを開始するようコードを編成する必要があることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="169ab-123">Note that because the **getAsync** method is asynchronous, if there are subsequent actions that depend on successfully getting the recipients, you should organize your code to start such actions only in the corresponding callback method when the asynchronous call has successfully completed.</span></span>




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


## <a name="set-recipients"></a><span data-ttu-id="169ab-124">受信者を設定する</span><span class="sxs-lookup"><span data-stu-id="169ab-124">Set recipients</span></span>


<span data-ttu-id="169ab-p107">このセクションでは、ユーザーが新規作成する予定やメッセージの受信者を設定するコード例を示しています。受信者を設定すると、既存の受信者が上書きされます。この例では、前述の新規作成フォームで受信者を取得する例と同様に、アドインが予定とメッセージの新規作成フォームでアクティブ化されることを想定しています。この例では、まず予定かメッセージのうち該当する方の受信者を表す適切なプロパティに対して非同期メソッド **Recipients.setAsync** を適用するために、新規作成するアイテムが予定かメッセージかを確認します。</span><span class="sxs-lookup"><span data-stu-id="169ab-p107">This section shows a code sample that sets the recipients of the appointment or message that is being composed by the user. Setting recipients overwrites any existing recipients. Similar to the previous example that gets recipients in a compose form, this example assumes that the add-in is activated in compose forms for appointments and messages. This example first verifies if the composed item is an appointment or message, so to apply the asynchronous method, **Recipients.setAsync**, on the appropriate properties that represent recipients of the appointment or message.</span></span>

<span data-ttu-id="169ab-129">**setAsync** を呼び出す際に、次に示すいずれかの形式で、_recipients_ パラメーターの入力引数として配列を指定します。</span><span class="sxs-lookup"><span data-stu-id="169ab-129">When calling **setAsync**, provide an array as input argument for the  _recipients_ parameter, in one of the following formats:</span></span>


- <span data-ttu-id="169ab-130">SMTP アドレスである文字列の配列。</span><span class="sxs-lookup"><span data-stu-id="169ab-130">An array of strings that are SMTP addresses.</span></span>
    
- <span data-ttu-id="169ab-131">辞書の配列。次のコード例に示されているように、それぞれ表示名と電子メール アドレスが含まれています。</span><span class="sxs-lookup"><span data-stu-id="169ab-131">An array of dictionaries, each containing a display name and email address, as shown in the following code sample.</span></span>
    
- <span data-ttu-id="169ab-132">**getAsync** メソッドによって返される配列のような、**EmailAddressDetails** オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="169ab-132">An array of **EmailAddressDetails** objects, similar to the one returned by the **getAsync** method.</span></span>
    
<span data-ttu-id="169ab-p108">オプションで **setAsync** メソッドに入力引数としてコールバック メソッドを提供し、受信者が正常に設定されることに依存するコードが、そのとおりになった場合に限り実行されるようにすることができます。オプションの _asyncContext_ パラメーターを使用してコールバック メソッドの引数を提供することもできます。コールバック メソッドを使用する場合は、_asyncResult_ 出力パラメーターにアクセスし、**AsyncResult** パラメーター オブジェクトの **status** プロパティや **error** プロパティを使用して非同期呼び出しの状態やエラー メッセージをチェックできます。</span><span class="sxs-lookup"><span data-stu-id="169ab-p108">You can optionally provide a callback method as an input argument to the **setAsync** method, to make sure any code that depends on successfully setting the recipients would execute only when that happens. You can also provide any arguments for the callback method using the optional _asyncContext_ parameter. If you use a callback method, you can access an _asyncResult_ output parameter, and use the **status** and **error** properties of the **AsyncResult** parameter object to check for status and any error messages of the asynchronous call.</span></span>




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


## <a name="add-recipients"></a><span data-ttu-id="169ab-136">受信者を追加する</span><span class="sxs-lookup"><span data-stu-id="169ab-136">Add recipients</span></span>


<span data-ttu-id="169ab-p109">予定やメッセージの既存の受信者を上書きしない場合は、**Recipients.setAsync** を使用する代わりに、**Recipients.addAsync** 非同期メソッドを使用して受信者を付加できます。**addAsync** の働きは **setAsync** と似ていて、_recipients_ 入力引数が必要です。オプションで、コールバック メソッドを指定し、asyncContext パラメーターを使用してコールバックの引数を提供できます。その後コールバック メソッドの **asyncResult** 出力パラメーターを使用して、非同期 _addAsync_ 呼び出しの状態、結果、エラーをチェックできます。次の例は、新規作成されるアイテムが予定かどうかチェックし、その予定に 2 人の必須の出席者を付加します。</span><span class="sxs-lookup"><span data-stu-id="169ab-p109">If you do not want to overwrite any existing recipients in an appointment or message, instead of using **Recipients.setAsync**, you can use the **Recipients.addAsync** asynchronous method to append recipients. **addAsync** works similarly as **setAsync** in that it requires a _recipients_ input argument. You can optionally provide a callback method, and any arguments for the callback using the asyncContext parameter. You can then check the status, result, and any error of the asynchronous **addAsync** call by using the _asyncResult_ output parameter of the callback method. The following example checks if the item being composed is an appointment, and appends two required attendees to the appointment.</span></span>


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


## <a name="see-also"></a><span data-ttu-id="169ab-142">関連項目</span><span class="sxs-lookup"><span data-stu-id="169ab-142">See also</span></span>

- [<span data-ttu-id="169ab-143">Outlook で新規作成フォームのアイテム データを取得および設定する</span><span class="sxs-lookup"><span data-stu-id="169ab-143">Get and set item data in a compose form in Outlook</span></span>](get-and-set-item-data-in-a-compose-form.md)    
- [<span data-ttu-id="169ab-144">閲覧または新規作成フォームの Outlook アイテム データを取得および設定する</span><span class="sxs-lookup"><span data-stu-id="169ab-144">Get and set Outlook item data in read or compose forms</span></span>](item-data.md)   
- [<span data-ttu-id="169ab-145">新規作成フォーム用の Outlook アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="169ab-145">Create Outlook add-ins for compose forms</span></span>](compose-scenario.md)    
- [<span data-ttu-id="169ab-146">Office アドインにおける非同期プログラミング</span><span class="sxs-lookup"><span data-stu-id="169ab-146">Asynchronous programming in Office Add-ins</span></span>](../develop/asynchronous-programming-in-office-add-ins.md)    
- [<span data-ttu-id="169ab-147">Outlook で予定またはメッセージを作成するときに件名を取得または設定する</span><span class="sxs-lookup"><span data-stu-id="169ab-147">Get or set the subject when composing an appointment or message in Outlook</span></span>](get-or-set-the-subject.md)    
- [<span data-ttu-id="169ab-148">Outlook で予定またはメッセージを作成するときに本文にデータを挿入する</span><span class="sxs-lookup"><span data-stu-id="169ab-148">Insert data in the body when composing an appointment or message in Outlook</span></span>](insert-data-in-the-body.md)    
- [<span data-ttu-id="169ab-149">Outlook で予定を作成するときに場所を取得または設定する</span><span class="sxs-lookup"><span data-stu-id="169ab-149">Get or set the location when composing an appointment in Outlook</span></span>](get-or-set-the-location-of-an-appointment.md) 
- [<span data-ttu-id="169ab-150">Outlook で予定を作成するときに時刻を取得または設定する</span><span class="sxs-lookup"><span data-stu-id="169ab-150">Get or set the time when composing an appointment in Outlook</span></span>](get-or-set-the-time-of-an-appointment.md)
    
