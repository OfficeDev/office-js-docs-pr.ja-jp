---
title: Outlook アドインで本文にデータを挿入する
description: Outlook アドインで、メッセージまたは予定の本文にデータを挿入する方法について説明します。
ms.date: 07/08/2022
ms.localizationpriority: medium
ms.openlocfilehash: 7319a3bb41d857fcae32ea118a3f3e60197bf751
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958328"
---
# <a name="insert-data-in-the-body-when-composing-an-appointment-or-message-in-outlook"></a>Outlook で予定またはメッセージを作成するときに本文にデータを挿入する

非同期メソッド ([Body.getAsync](/javascript/api/outlook/office.body#outlook-office-body-getasync-member(1))、[Body.getTypeAsync](/javascript/api/outlook/office.body#outlook-office-body-gettypeasync-member(1))、[Body.prependAsync](/javascript/api/outlook/office.body#outlook-office-body-prependasync-member(1))、[Body.setAsync](/javascript/api/outlook/office.body#outlook-office-body-setasync-member(1)) および [Body.setSelectedDataAsync](/javascript/api/outlook/office.body#outlook-office-body-setselecteddataasync-member(1))) を使用して、本文の種類を取得し、ユーザーが作成している予定またはメッセージのアイテムの本文にデータを挿入することができます。これらの非同期メソッドは新規作成アドインでのみ使用できます。これらのメソッドを使用する場合は、Outlook が新規作成フォーム内でアドインをアクティブ化できるようにアドイン マニフェストが適切にセットアップされていることを確認してください。この手順については、「[新規作成フォーム用の Outlook アドインを作成する](compose-scenario.md)」を参照してください。

Outlook では、メッセージはテキスト形式、HTML 形式、またはリッチ テキスト形式 (RTF) で作成でき、予定は HTML 形式で作成できます。 挿入する前に、必ず **getTypeAsync** を呼び出して、サポートされている項目の形式を確認する必要があります。追加の手順を実行する必要がある場合があります。 **getTypeAsync** が返す値は、元の項目の形式と、HTML 形式での編集に対するデバイス オペレーティング システムとアプリケーションのサポート (1) によって異なります。 次に、次の表に示すように、(2) に応じて **prependAsync** または **setSelectedDataAsync** の _coercionType_ パラメーターを設定してデータを挿入します。 引数を指定しない場合、 **prependAsync** および **setSelectedDataAsync** は挿入するデータがテキスト形式であると想定します。

|挿入するデータ|getTypeAsync によって返されるアイテム形式|使用する coercionType|
|:-----|:-----|:-----|
|テキスト|テキスト (1)|テキスト|
|HTML|テキスト (1)|テキスト (2)|
|テキスト|HTML|テキスト/HTML|
|HTML|HTML |HTML|

1. タブレットやスマートフォンでは、オペレーティング システムまたはアプリケーションが HTML 形式で最初に作成されたアイテムの編集をサポートしていない場合、 **getTypeAsync** は **Office.MailboxEnums.BodyType.Text** を返します。

1. 挿入するデータが HTML であり **、getTypeAsync** によってその項目のテキスト型が返される場合は、データをテキストとして再構成し、 **Office.MailboxEnums.BodyType.Text** を _coercionType_ として挿入します。 テキスト強制型の HTML データを単純に挿入する場合、アプリケーションは HTML タグをテキストとして表示します。 **Office.MailboxEnums.BodyType.Html** を _coercionType_ として HTML データに挿入しようとすると、エラーが発生します。

_CoercionType_ に加えて、Office JavaScript API のほとんどの非同期メソッドと同様に、**getTypeAsync**、**prependAsync**、**setSelectedDataAsync** は、他のオプションの入力パラメーターを受け取ります。 これらのオプションの入力パラメーターの指定について詳しくは、「 [Office アドインにおける非同期プログラミング](../develop/asynchronous-programming-in-office-add-ins.md#pass-optional-parameters-inline)」の「 [オプションのパラメーターを非同期メソッドに渡す](../develop/asynchronous-programming-in-office-add-ins.md)」を参照してください。

## <a name="insert-data-at-the-current-cursor-position"></a>現在のカーソル位置にデータを挿入する

ここでは、作成中のアイテムの本文タイプを **getTypeAsync** を使用して検査してから、**setSelectedDataAsync** を使用して現在のカーソル位置にデータを挿入するサンプル コードを示します。

コールバック関数とオプションの入力パラメーターを **getTypeAsync** に渡し、状態と結果を  _asyncResult_ 出力パラメーターで取得できます。 メソッドが成功した場合は、 [AsyncResult.value](/javascript/api/office/office.asyncresult#office-office-asyncresult-value-member) プロパティ ("text" または "html" のいずれか) で項目本文の型を取得できます。

**setSelectedDataAsync** に入力パラメーターとしてデータ文字列を渡す必要があります。 アイテム本文のタイプに応じて、このデータ文字列はテキスト形式または HTML 形式で指定できます。 前述したように、挿入するデータのタイプを _coercionType_ パラメーターで指定できます。 また、コールバック関数とそのパラメーターを省略可能な入力パラメーターとして指定することもできます。

ユーザーがアイテム本文にカーソルを置いていない場合、**setSelectedDataAsync** はデータを本文の先頭に挿入します。ユーザーがアイテム本文のテキストを選択した場合、**setSelectedDataAsync** は選択されたテキストを指定されたデータに置き換えます。ユーザーがカーソル位置の変更とアイテムの作成を同時に行う場合、**setSelectedDataAsync** が失敗する可能性があることに注意してください。1 回で挿入できる文字の最大数は、1,000,000 文字です。

このサンプル コードは、以下に示すように、アドイン マニフェストのルールが、予定またはメッセージの新規作成フォームでアドインをアクティブにすることを想定しています。

```XML
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit"/>
</Rule>
```

```js
let item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready method.
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

## <a name="insert-data-at-the-beginning-of-the-item-body"></a>アイテムの本文の先頭にデータを挿入する

別の方法として、**prependAsync** を使用して現在のカーソル位置にかかわらず、データをアイテム本文の先頭に挿入することもできます。挿入の位置が異なることを除けば、**prependAsync** と **setSelectedDataAsync** の動作は同じです。

- メッセージ本文で HTML データの先頭に HTML データを追加する場合は、最初にメッセージ本文の種類を確認して、テキスト形式のメッセージに HTML データを追加しないようにする必要があります。

- **prependAsync** の入力パラメーターとして、テキストまたは HTML 形式のデータ文字列、および必要に応じて挿入するデータの形式、コールバック関数、およびそのパラメーターのいずれかを指定します。

- 同時に先頭に付加できる文字の最大数は、1,000,000 文字です。

以下の JavaScript コードは、予定およびメッセージの新規作成フォーム内でアクティブ化されるサンプル アドインの一部です。このサンプルは、**getTypeAsync** を呼び出して、アイテム本文の種類を検査し、アイテムが予定または HTML メッセージの場合にはアイテム本文の先頭に HTML データを付加し、それ以外の場合にはテキスト形式のデータを挿入します。

```js
let item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready method.
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

## <a name="see-also"></a>関連項目

- [Outlook で新規作成フォームのアイテム データを取得および設定する](get-and-set-item-data-in-a-compose-form.md)
- [閲覧または新規作成フォームの Outlook アイテム データを取得および設定する](item-data.md)
- [新規作成フォーム用の Outlook アドインを作成する](compose-scenario.md)
- [Office アドインにおける非同期プログラミング](../develop/asynchronous-programming-in-office-add-ins.md)
- [Outlook の予定またはメッセージを作成するときに受信者を取得、設定、追加する](get-set-or-add-recipients.md)  
- [Outlook で予定またはメッセージを作成するときに件名を取得または設定する](get-or-set-the-subject.md)  
- [Outlook で予定を作成するときに場所を取得または設定する](get-or-set-the-location-of-an-appointment.md)
- [Outlook で予定を作成するときに時刻を取得または設定する](get-or-set-the-time-of-an-appointment.md)
