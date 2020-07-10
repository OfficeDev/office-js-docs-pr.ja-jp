---
title: ホストページからダイアログボックスにデータとメッセージを渡す
description: MessageChild および DialogParentMessageReceived Api を使用して、ホストページからダイアログにデータを渡す方法について説明します。
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: 05220fa4cecad4fe412a5590605f774f92ef8f61
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093575"
---
# <a name="passing-data-and-messages-to-a-dialog-box-from-its-host-page-preview"></a>ホストページからダイアログボックスにデータとメッセージを渡す (プレビュー)

アドインでは、 [dialog](/javascript/api/office/office.dialog)オブジェクトの[messageChild](/javascript/api/office/office.dialog#messagechild-message-)メソッドを使用して、[ホストページ](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page)からダイアログボックスにメッセージを送信できます。

> [!Important]
>
> - この記事で説明する Api はプレビュー段階です。 開発者は実験を行うことができます。ただし、運用アドインでは使用しないでください。 この API がリリースされるまでは、「次の操作を実行するには」で説明されている方法を使用して、運用アドインの[ダイアログボックスに情報を渡し](dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box)ます。
> - この記事に記載されている Api には、Microsoft 365 のサブスクリプションが必要です。 Insider チャネルからの最新の月次バージョンとビルドを使ってください。 このバージョンを入手するには、Office Insider への参加が必要です。 詳細については、「[Office Insider になる](https://insider.office.com)」を参照してください。 ビルドが生産半期チャネルに graduates されている場合、そのビルドではプレビュー機能のサポートがオフになっていることに注意してください。
> - プレビューの初期段階では、Excel、PowerPoint、Word で Api がサポートされています。ただし、Outlook には含まれません。
>
> [!INCLUDE [Information about using preview APIs](../includes/using-preview-apis.md)]

## <a name="use-messagechild-from-the-host-page"></a>`messageChild()`ホストページからの使用

Office ダイアログ API を呼び出してダイアログボックスを開くと、 [dialog](/javascript/api/office/office.dialog)オブジェクトが返されます。 オブジェクトは他のメソッドによって参照されるため、通常は[Displaydialogasync](/javascript/api/office/office.ui#displaydialogasync-startaddress--callback-)メソッドよりも広いスコープがある変数に割り当てる必要があります。 例を次に示します。

```javascript
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
    function (asyncResult) {
        dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
);

function processMessage(arg) {
    dialog.close();

  // message processing code goes here;

}
```

このオブジェクトには、 `Dialog` すべての文字列または文字列データをダイアログボックスに送信する[messageChild](/javascript/api/office/office.dialog#messagechild-message-)メソッドがあります。 これにより `DialogParentMessageReceived` 、ダイアログボックスでイベントが発生します。 コードでは、次のセクションに示すように、このイベントを処理する必要があります。

ダイアログの UI が現在アクティブなワークシートと関連付けられ、他のワークシートを基準としたワークシートの相対位置になるシナリオを考えてみます。 次の例では、 `sheetPropertiesChanged` Excel ワークシートのプロパティをダイアログボックスに送信します。 この例では、現在のワークシートの名前は "My Sheet" で、ブックの2番目のシートです。 データは、文字列のオブジェクトにカプセル化されるので、に渡すことができ `messageChild` ます。

```javascript
function sheetPropertiesChanged() {
    var messageToDialog = JSON.stringify({
                               name: "My Sheet",
                               position: 2
                           });

    dialog.messageChild(messageToDialog);
}
```

## <a name="handle-dialogparentmessagereceived-in-the-dialog-box"></a>ダイアログボックスで DialogParentMessageReceived を処理する

ダイアログボックスの JavaScript で、イベントのハンドラーを `DialogParentMessageReceived` [UI. Addhandler async](/javascript/api/office/office.ui#addhandlerasync-eventtype--handler--options--callback-)メソッドに登録します。 これは、通常、 [tialize メソッドまたは Office.iniメソッド](initialize-add-in.md)で行われます。 例を次に示します。

```javascript
Office.onReady()
    .then(function() {
        Office.context.ui.addHandlerAsync(
            Office.EventType.DialogParentMessageReceived,
            onMessageFromParent);
    });
```

その後、ハンドラーを定義し `onMessageFromParent` ます。 次のコードでは、前のセクションの例を続行します。 Office によってハンドラーに引数が渡され、 `message` 引数オブジェクトのプロパティにホストページの文字列が含まれていることに注意してください。 この例では、メッセージはオブジェクトに再変換、jQuery を使用して、新しいワークシート名に一致するダイアログのトップの見出しを設定しています。

```javascript
function onMessageFromParent(event) {
    var messageFromParent = JSON.parse(event.message);
    $('h1').text(messageFromParent.name);
}
```

ハンドラーが適切に登録されていることを確認することをお勧めします。 これを行うには、 `addHandlerAsync` ハンドラーの登録が完了したときに実行されるメソッドにコールバックを渡します。 ハンドラーが正常に登録されなかった場合は、ハンドラーを使用して、エラーを記録または表示します。 次に例を示します。 ここで `reportError` は、エラーを記録または表示する関数であることに注意してください。

```javascript
Office.onReady()
    .then(function() {
        Office.context.ui.addHandlerAsync(
            Office.EventType.DialogParentMessageReceived,
            onMessageFromParent,
            onRegisterMessageComplete);
    });

function onRegisterMessageComplete(asyncResult) {
    if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        reportError(asyncResult.error.message);
    }
}
```

## <a name="conditional-messaging"></a>条件付きのメッセージング

ホストページから複数の呼び出しを行うことはできます `messageChild` が、イベントのダイアログボックスにはハンドラーが1つしかないため、 `DialogParentMessageReceived` ハンドラーは異なるメッセージを区別するために条件付きロジックを使用する必要があります。 [条件付き](dialog-api-in-office-add-ins.md#conditional-messaging)メッセージの説明に従って、ダイアログボックスがホストページにメッセージを送信しているときに、条件付きメッセージを構造化する方法で、これを正確に行うことができます。
