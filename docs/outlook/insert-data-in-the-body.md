---
title: Outlook アドインで本文にデータを挿入する
description: Outlook アドインで、メッセージまたは予定の本文にデータを挿入する方法について説明します。
ms.date: 04/15/2019
localization_priority: Normal
ms.openlocfilehash: 0e875619520ee309dec97b2db60ed49c29b2a463
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293871"
---
# <a name="insert-data-in-the-body-when-composing-an-appointment-or-message-in-outlook"></a><span data-ttu-id="81d75-103">Outlook で予定またはメッセージを作成するときに本文にデータを挿入する</span><span class="sxs-lookup"><span data-stu-id="81d75-103">Insert data in the body when composing an appointment or message in Outlook</span></span>

<span data-ttu-id="81d75-p101">非同期メソッド ([Body.getAsync](/javascript/api/outlook/office.Body#getasync-coerciontype--options--callback-)、[Body.getTypeAsync](/javascript/api/outlook/office.Body#gettypeasync-options--callback-)、[Body.prependAsync](/javascript/api/outlook/office.Body#prependasync-data--options--callback-)、[Body.setAsync](/javascript/api/outlook/office.Body#setasync-data--options--callback-) および [Body.setSelectedDataAsync](/javascript/api/outlook/office.Body#setselecteddataasync-data--options--callback-)) を使用して、本文の種類を取得し、ユーザーが作成している予定またはメッセージのアイテムの本文にデータを挿入することができます。これらの非同期メソッドは新規作成アドインでのみ使用できます。これらのメソッドを使用する場合は、Outlook が新規作成フォーム内でアドインをアクティブ化できるようにアドイン マニフェストが適切にセットアップされていることを確認してください。この手順については、「[新規作成フォーム用の Outlook アドインを作成する](compose-scenario.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="81d75-p101">You can use the asynchronous methods ([Body.getAsync](/javascript/api/outlook/office.Body#getasync-coerciontype--options--callback-), [Body.getTypeAsync](/javascript/api/outlook/office.Body#gettypeasync-options--callback-), [Body.prependAsync](/javascript/api/outlook/office.Body#prependasync-data--options--callback-), [Body.setAsync](/javascript/api/outlook/office.Body#setasync-data--options--callback-) and [Body.setSelectedDataAsync](/javascript/api/outlook/office.Body#setselecteddataasync-data--options--callback-)) to get the body type and insert data in the body of an appointment or message item that the user is composing. These asynchronous methods are available to only compose add-ins. To use these methods, make sure you have set up the add-in manifest appropriately so that Outlook activates your add-in in compose forms, as described in [Create Outlook add-ins for compose forms](compose-scenario.md).</span></span>

<span data-ttu-id="81d75-p102">Outlook では、ユーザーはテキスト、HTML、またはリッチテキスト形式 (RTF) でメッセージを作成することができ、HTML 形式で予定を作成することができます。挿入する前に、必ず**getTypeAsync**を呼び出して、サポートされているアイテム形式を確認する必要があります。これには、追加の手順を実行する必要があります。**GetTypeAsync**が返す値は、元のアイテムの形式、および HTML 形式 (1) で編集するデバイスのオペレーティングシステムとアプリケーションのサポートによって異なります。その後、次の表に示すように、 **prependAsync**または**Setselecteddataasync**の_coercionType_パラメーターを設定して、データを挿入します。引数を指定しない場合、 **prependAsync**と**Setselecteddataasync**は挿入するデータがテキスト形式であると仮定します。</span><span class="sxs-lookup"><span data-stu-id="81d75-p102">In Outlook, a user can create a message in text, HTML, or Rich Text Format (RTF), and can create an appointment in HTML format. Before inserting, you should always first verify the supported item format by calling **getTypeAsync**, as you may need to take additional steps. The value that **getTypeAsync** returns depends on the original item format, as well as the support of the device operating system and application to editing in HTML format (1). Then set the  _coercionType_ parameter of **prependAsync** or **setSelectedDataAsync** accordingly (2) to insert the data, as shown in the following table. If you don't specify an argument, **prependAsync** and **setSelectedDataAsync** assume the data to insert is in text format.</span></span>

<br/>

|<span data-ttu-id="81d75-111">**挿入するデータ**</span><span class="sxs-lookup"><span data-stu-id="81d75-111">**Data to insert**</span></span>|<span data-ttu-id="81d75-112">**getTypeAsync によって返されるアイテム形式**</span><span class="sxs-lookup"><span data-stu-id="81d75-112">**Item format returned by getTypeAsync**</span></span>|<span data-ttu-id="81d75-113">**使用する coercionType**</span><span class="sxs-lookup"><span data-stu-id="81d75-113">**Use this coercionType**</span></span>|
|:-----|:-----|:-----|
|<span data-ttu-id="81d75-114">テキスト</span><span class="sxs-lookup"><span data-stu-id="81d75-114">Text</span></span>|<span data-ttu-id="81d75-115">テキスト (1)</span><span class="sxs-lookup"><span data-stu-id="81d75-115">Text (1)</span></span>|<span data-ttu-id="81d75-116">テキスト</span><span class="sxs-lookup"><span data-stu-id="81d75-116">Text</span></span>|
|<span data-ttu-id="81d75-117">HTML</span><span class="sxs-lookup"><span data-stu-id="81d75-117">HTML</span></span>|<span data-ttu-id="81d75-118">テキスト (1)</span><span class="sxs-lookup"><span data-stu-id="81d75-118">Text (1)</span></span>|<span data-ttu-id="81d75-119">テキスト (2)</span><span class="sxs-lookup"><span data-stu-id="81d75-119">Text (2)</span></span>|
|<span data-ttu-id="81d75-120">テキスト</span><span class="sxs-lookup"><span data-stu-id="81d75-120">Text</span></span>|<span data-ttu-id="81d75-121">HTML</span><span class="sxs-lookup"><span data-stu-id="81d75-121">HTML</span></span>|<span data-ttu-id="81d75-122">テキスト/HTML</span><span class="sxs-lookup"><span data-stu-id="81d75-122">Text/HTML</span></span>|
|<span data-ttu-id="81d75-123">HTML</span><span class="sxs-lookup"><span data-stu-id="81d75-123">HTML</span></span>|<span data-ttu-id="81d75-124">HTML</span><span class="sxs-lookup"><span data-stu-id="81d75-124">HTML</span></span> |<span data-ttu-id="81d75-125">HTML</span><span class="sxs-lookup"><span data-stu-id="81d75-125">HTML</span></span>|

1.  <span data-ttu-id="81d75-126">タブレットとスマートフォンでは、 **getTypeAsync** は **MailboxEnums の種類のテキスト** を返します。これは、オペレーティングシステムまたはアプリケーションが html 形式で作成されたアイテムの編集をサポートしていない場合です。</span><span class="sxs-lookup"><span data-stu-id="81d75-126">On tablets and smartphones, **getTypeAsync** returns **Office.MailboxEnums.BodyType.Text** if the operating system or application does not support editing an item, which was originally created in HTML, in HTML format.</span></span>

2.  <span data-ttu-id="81d75-p103">挿入するデータが HTML で、 **getTypeAsync** がそのアイテムのテキストタイプを返す場合は、データをテキストとして再編成し、 **MailboxEnums** として _coercionType_として挿入します。テキストの強制型変換を使用して HTML データを挿入するだけの場合、アプリケーションでは HTML タグがテキストとして表示されます。 **Office.MailboxEnums.BodyType.Html** を _coercionType_として HTML データを挿入しようとすると、エラーが表示されます。</span><span class="sxs-lookup"><span data-stu-id="81d75-p103">If your data to insert is HTML and **getTypeAsync** returns a text type for that item, reorganize your data as text and insert it with **Office.MailboxEnums.BodyType.Text** as _coercionType_. If you simply insert the HTML data with a text coercion type, the application would display the HTML tags as text. If you attempt to insert the HTML data with **Office.MailboxEnums.BodyType.Html** as _coercionType_, you will get an error.</span></span>

<span data-ttu-id="81d75-p104">_CoercionType_に加えて、OFFICE JavaScript API のほとんどの非同期メソッドと同様に、 **getTypeAsync**、 **PrependAsync** 、および**setselecteddataasync**は、他のオプションの入力パラメーターを受け取ります。これらのオプション入力パラメーターの指定の詳細については、「 [Office アドインで非同期プログラミング](../develop/asynchronous-programming-in-office-add-ins.md)を使用した[非同期メソッドへのオプションパラメーターの引き渡し](../develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-inline)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="81d75-p104">In addition to  _coercionType_, as with most asynchronous methods in the Office JavaScript API, **getTypeAsync**, **prependAsync** and **setSelectedDataAsync** take other optional input parameters. For more information about specifying these optional input parameters, see [passing optional parameters to asynchronous methods](../develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-inline) in [Asynchronous programming in Office Add-ins](../develop/asynchronous-programming-in-office-add-ins.md).</span></span>


## <a name="insert-data-at-the-current-cursor-position"></a><span data-ttu-id="81d75-132">現在のカーソル位置にデータを挿入する</span><span class="sxs-lookup"><span data-stu-id="81d75-132">Insert data at the current cursor position</span></span>


<span data-ttu-id="81d75-133">ここでは、作成中のアイテムの本文タイプを **getTypeAsync** を使用して検査してから、**setSelectedDataAsync** を使用して現在のカーソル位置にデータを挿入するサンプル コードを示します。</span><span class="sxs-lookup"><span data-stu-id="81d75-133">This section shows a code sample that uses **getTypeAsync** to verify the body type of the item that is being composed, and then uses **setSelectedDataAsync** to insert data in the current cursor location.</span></span>

<span data-ttu-id="81d75-p105">コールバック メソッドとオプションの入力パラメーターを **getTypeAsync** に渡し、ステータスと結果を _asyncResult_ 出力パラメーターで受け取ることができます。メソッドが成功した場合、アイテム本文のタイプを [AsyncResult.value](/javascript/api/office/office.asyncresult#value) プロパティで受け取ることができます。その値は、"text" または "html" です。</span><span class="sxs-lookup"><span data-stu-id="81d75-p105">You can pass a callback method and optional input parameters to **getTypeAsync**, and get any status and results in the  _asyncResult_ output parameter. If the method succeeds, you can get the type of the item body in the [AsyncResult.value](/javascript/api/office/office.asyncresult#value) property, which is either "text" or "html".</span></span>

<span data-ttu-id="81d75-p106">**setSelectedDataAsync** への入力パラメーターとして、データ文字列を渡す必要があります。アイテム本文のタイプに応じて、このデータ文字列はテキスト形式または HTML 形式で指定できます。前述したように、挿入するデータのタイプを _coercionType_ パラメーターで指定できます。また、コールバック メソッドとそのパラメーターをオプションの入力パラメーターとして指定できます。</span><span class="sxs-lookup"><span data-stu-id="81d75-p106">You must pass a data string as an input parameter to **setSelectedDataAsync**. Depending on the type of the item body, you can specify this data string in text or HTML format accordingly. As mentioned above, you can optionally specify the type of the data to be inserted in the  _coercionType_ parameter. In addition, you can provide a callback method and any of its parameters as optional input parameters.</span></span>

<span data-ttu-id="81d75-p107">ユーザーがアイテム本文にカーソルを置いていない場合、**setSelectedDataAsync** はデータを本文の先頭に挿入します。ユーザーがアイテム本文のテキストを選択した場合、**setSelectedDataAsync** は選択されたテキストを指定されたデータに置き換えます。ユーザーがカーソル位置の変更とアイテムの作成を同時に行う場合、**setSelectedDataAsync** が失敗する可能性があることに注意してください。1 回で挿入できる文字の最大数は、1,000,000 文字です。</span><span class="sxs-lookup"><span data-stu-id="81d75-p107">If the user hasn't placed the cursor in the item body, **setSelectedDataAsync** inserts the data at the top of the body. If the user has selected text in the item body, **setSelectedDataAsync** replaces the selected text by the data you specify. Note that **setSelectedDataAsync** can fail if the user is simultaneously changing the cursor position while composing the item. The maximum number of characters you can insert at one time is 1,000,000 characters.</span></span>

<span data-ttu-id="81d75-144">このサンプル コードは、以下に示すように、アドイン マニフェストのルールが、予定またはメッセージの新規作成フォームでアドインをアクティブにすることを想定しています。</span><span class="sxs-lookup"><span data-stu-id="81d75-144">This code sample assumes a rule in the add-in manifest that activates the add-in in a compose form for an appointment or message, as shown below.</span></span>




```XML
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit"/>
</Rule>

```




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set data in the body of the composed item.
        setItemBody();
    });
}


// Get the body type of the composed item, and set data in 
// in the appropriate data type in the item body.
function setItemBody() {
    item.body.getTypeAsync(
        function (result) {
            if (result.status == Office.AsyncResultStatus.Failed){
                write(result.error.message);
            }
            else {
                // Successfully got the type of item body.
                // Set data of the appropriate type in body.
                if (result.value == Office.MailboxEnums.BodyType.Html) {
                    // Body is of HTML type.
                    // Specify HTML in the coercionType parameter
                    // of setSelectedDataAsync.
                    item.body.setSelectedDataAsync(
                        '<b> Kindly note we now open 7 days a week.</b>',
                        { coercionType: Office.CoercionType.Html, 
                        asyncContext: { var3: 1, var4: 2 } },
                        function (asyncResult) {
                            if (asyncResult.status == 
                                Office.AsyncResultStatus.Failed){
                                write(asyncResult.error.message);
                            }
                            else {
                                // Successfully set data in item body.
                                // Do whatever appropriate for your scenario,
                                // using the arguments var3 and var4 as applicable.
                            }
                        });
                }
                else {
                    // Body is of text type. 
                    item.body.setSelectedDataAsync(
                        ' Kindly note we now open 7 days a week.',
                        { coercionType: Office.CoercionType.Text, 
                            asyncContext: { var3: 1, var4: 2 } },
                        function (asyncResult) {
                            if (asyncResult.status == 
                                Office.AsyncResultStatus.Failed){
                                write(asyncResult.error.message);
                            }
                            else {
                                // Successfully set data in item body.
                                // Do whatever appropriate for your scenario,
                                // using the arguments var3 and var4 as applicable.
                            }
                         });
                }
            }
        });

}

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## <a name="insert-data-at-the-beginning-of-the-item-body"></a><span data-ttu-id="81d75-145">アイテムの本文の先頭にデータを挿入する</span><span class="sxs-lookup"><span data-stu-id="81d75-145">Insert data at the beginning of the item body</span></span>


<span data-ttu-id="81d75-p108">別の方法として、**prependAsync** を使用して現在のカーソル位置にかかわらず、データをアイテム本文の先頭に挿入することもできます。挿入の位置が異なることを除けば、**prependAsync** と **setSelectedDataAsync** の動作は同じです。</span><span class="sxs-lookup"><span data-stu-id="81d75-p108">Alternatively, you can use **prependAsync** to insert data at the beginning of the item body and disregard the current cursor location. Other than the point of insertion, **prependAsync** and **setSelectedDataAsync** behave in similar ways:</span></span>


- <span data-ttu-id="81d75-148">メッセージ本文の先頭に HTML データを付加する場合、テキスト形式のメッセージの先頭に HTML データを付加することがないように、まずメッセージ本文のタイプを調べる必要があります。</span><span class="sxs-lookup"><span data-stu-id="81d75-148">If you are prepending HTML data in a message body, you should first check for the type of the message body to avoid prepending HTML data to a message in text format.</span></span>
    
- <span data-ttu-id="81d75-149">**prependAsync** への入力パラメーターとして、テキスト形式または HTML 形式のデータ文字列、挿入されるデータの形式 (オプション)、コールバック メソッド、およびそのパラメーター (パラメーターがある場合) を指定します。</span><span class="sxs-lookup"><span data-stu-id="81d75-149">Provide the following as input parameters to **prependAsync**: a data string in either text or HTML format, and optionally the format of the data to be inserted, a callback method and any of its parameters.</span></span>
    
- <span data-ttu-id="81d75-150">同時に先頭に付加できる文字の最大数は、1,000,000 文字です。</span><span class="sxs-lookup"><span data-stu-id="81d75-150">The maximum number of characters you can prepend at one time is 1,000,000 characters.</span></span>
    
<span data-ttu-id="81d75-p109">以下の JavaScript コードは、予定およびメッセージの新規作成フォーム内でアクティブ化されるサンプル アドインの一部です。このサンプルは、**getTypeAsync** を呼び出して、アイテム本文の種類を検査し、アイテムが予定または HTML メッセージの場合にはアイテム本文の先頭に HTML データを付加し、それ以外の場合にはテキスト形式のデータを挿入します。</span><span class="sxs-lookup"><span data-stu-id="81d75-p109">The following JavaScript code is part of a sample add-in that is activated in compose forms of appointments and messages. The sample calls **getTypeAsync** to verify the type of the item body, inserts HTML data to the top of the item body if the item is an appointment or HTML message, otherwise inserts the data in text format.</span></span>




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Insert data in the top of the body of the composed 
        // item.
        prependItemBody();
    });
}

// Get the body type of the composed item, and prepend data  
// in the appropriate data type in the item body.
function prependItemBody() {
    item.body.getTypeAsync(
        function (result) {
            if (result.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully got the type of item body.
                // Prepend data of the appropriate type in body.
                if (result.value == Office.MailboxEnums.BodyType.Html) {
                    // Body is of HTML type.
                    // Specify HTML in the coercionType parameter
                    // of prependAsync.
                    item.body.prependAsync(
                        '<b>Greetings!</b>',
                        { coercionType: Office.CoercionType.Html, 
                        asyncContext: { var3: 1, var4: 2 } },
                        function (asyncResult) {
                            if (asyncResult.status == 
                                Office.AsyncResultStatus.Failed){
                                write(asyncResult.error.message);
                            }
                            else {
                                // Successfully prepended data in item body.
                                // Do whatever appropriate for your scenario,
                                // using the arguments var3 and var4 as applicable.
                            }
                        });
                }
                else {
                    // Body is of text type. 
                    item.body.prependAsync(
                        'Greetings!',
                        { coercionType: Office.CoercionType.Text, 
                            asyncContext: { var3: 1, var4: 2 } },
                        function (asyncResult) {
                            if (asyncResult.status == 
                                Office.AsyncResultStatus.Failed){
                                write(asyncResult.error.message);
                            }
                            else {
                                // Successfully prepended data in item body.
                                // Do whatever appropriate for your scenario,
                                // using the arguments var3 and var4 as applicable.
                            }
                         });
                }
            }
        });

}

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## <a name="see-also"></a><span data-ttu-id="81d75-153">関連項目</span><span class="sxs-lookup"><span data-stu-id="81d75-153">See also</span></span>

- [<span data-ttu-id="81d75-154">Outlook で新規作成フォームのアイテム データを取得および設定する</span><span class="sxs-lookup"><span data-stu-id="81d75-154">Get and set item data in a compose form in Outlook</span></span>](get-and-set-item-data-in-a-compose-form.md)    
- [<span data-ttu-id="81d75-155">閲覧または新規作成フォームの Outlook アイテム データを取得および設定する</span><span class="sxs-lookup"><span data-stu-id="81d75-155">Get and set Outlook item data in read or compose forms</span></span>](item-data.md)    
- [<span data-ttu-id="81d75-156">新規作成フォーム用の Outlook アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="81d75-156">Create Outlook add-ins for compose forms</span></span>](compose-scenario.md)    
- [<span data-ttu-id="81d75-157">Office アドインにおける非同期プログラミング</span><span class="sxs-lookup"><span data-stu-id="81d75-157">Asynchronous programming in Office Add-ins</span></span>](../develop/asynchronous-programming-in-office-add-ins.md)    
- [<span data-ttu-id="81d75-158">Outlook の予定またはメッセージを作成するときに受信者を取得、設定、追加する</span><span class="sxs-lookup"><span data-stu-id="81d75-158">Get, set, or add recipients when composing an appointment or message in Outlook</span></span>](get-set-or-add-recipients.md)  
- [<span data-ttu-id="81d75-159">Outlook で予定またはメッセージを作成するときに件名を取得または設定する</span><span class="sxs-lookup"><span data-stu-id="81d75-159">Get or set the subject when composing an appointment or message in Outlook</span></span>](get-or-set-the-subject.md)  
- [<span data-ttu-id="81d75-160">Outlook で予定を作成するときに場所を取得または設定する</span><span class="sxs-lookup"><span data-stu-id="81d75-160">Get or set the location when composing an appointment in Outlook</span></span>](get-or-set-the-location-of-an-appointment.md) 
- [<span data-ttu-id="81d75-161">Outlook で予定を作成するときに時刻を取得または設定する</span><span class="sxs-lookup"><span data-stu-id="81d75-161">Get or set the time when composing an appointment in Outlook</span></span>](get-or-set-the-time-of-an-appointment.md)
    
