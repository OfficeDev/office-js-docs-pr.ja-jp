---
title: Office ダイアログ ボックスでのエラーとイベントの処理
description: Office ダイアログボックスを開いて使用するときに発生するエラーをトラップして処理する方法について説明します。
ms.date: 01/29/2020
localization_priority: Normal
ms.openlocfilehash: d83d5c4627f68c3f4b1c196cf543d01bf981abbe
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608175"
---
# <a name="handling-errors-and-events-in-the-office-dialog-box"></a>Office ダイアログ ボックスでのエラーとイベントの処理

この記事では、ダイアログボックスを開くときにエラーをトラップして処理する方法と、ダイアログボックス内で発生するエラーについて説明します。

> [!NOTE]
> この記事では、「office[アドインで office ダイアログ api を使用](dialog-api-in-office-add-ins.md)する」で説明されている OFFICE ダイアログ api の使用についての基本事項を presupposes しています。
> 
> 「 [Office ダイアログ API のベストプラクティスとルール](dialog-best-practices.md)」も参照してください。

コードでイベントの 2 つのカテゴリを処理する必要があります。

- ダイアログ ボックスを作成できないために `displayDialogAsync` の呼び出しによって返されるエラー。
- ダイアログボックス内のエラーおよびその他のイベント。

## <a name="errors-from-displaydialogasync"></a>displayDialogAsync のエラー

一般的なプラットフォームおよびシステムのエラーに加えて、4つのエラーが呼び出しに固有のものです `displayDialogAsync` 。

|コード番号|意味|
|:-----|:-----|
|12004|`displayDialogAsync` に渡される URL のドメインは信頼されていません。ドメインは、ホスト ページと同じドメインにある必要があります (プロトコルとポート番号を含む)。|
|12005|`displayDialogAsync` に渡される URL には HTTP プロトコルを使用します。 HTTPS が必要です。 (一部のバージョンの Office では、12005で返されるエラーメッセージのテキストは、12004で返されるものと同じです。)|
|<span id="12007">12007</span><!-- The span is needed because office-js-helpers has an error message that links to this table row. -->|ダイアログ ボックスは、このホスト ウィンドウで既に開いています。作業ウィンドウなどのホスト ウィンドウで一度に開けるダイアログ ボックスは 1 つだけです。|
|12009|ダイアログ ボックスを無視するようにユーザーが選択しました。 このエラーは、web 上の Office で発生する可能性があります。ユーザーは、アドインによるダイアログボックスの表示を許可しないことを選択できます。 詳細については、「 [web 上の Office を使用したポップアップブロックの処理](dialog-best-practices.md#handling-pop-up-blockers-with-office-on-the-web)」を参照してください。|

`displayDialogAsync`が呼び出されると、 [AsyncResult](/javascript/api/office/office.asyncresult)オブジェクトをコールバック関数に渡します。 呼び出しが成功すると、ダイアログボックスが開き、 `value` オブジェクトのプロパティ `AsyncResult` は[dialog](/javascript/api/office/office.dialog)オブジェクトになります。 この例については、「[送信情報をダイアログボックスからホストページに送信する](dialog-api-in-office-add-ins.md#send-information-from-the-dialog-box-to-the-host-page)」を参照してください。 呼び出しが失敗すると、 `displayDialogAsync` ダイアログボックスは作成されず、 `status` オブジェクトのプロパティ `AsyncResult` がに設定され、 `Office.AsyncResultStatus.Failed` `error` オブジェクトのプロパティが設定されます。 をテストして、 `status` エラーが発生したときに応答するコールバックを常に提供する必要があります。 コード番号に関係なくエラーメッセージを報告する例については、次のコードを参照してください。 ( `showNotification` この記事で定義されていない関数は、エラーを表示またはログ記録します。 アドイン内でこの関数を実装する方法の例については、「 [Office アドインダイアログ API の例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)」を参照してください。

```js
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        showNotification(asyncResult.error.code = ": " + asyncResult.error.message);
    } else {
        dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
});
```

## <a name="errors-and-events-in-the-dialog-box"></a>ダイアログボックスのエラーとイベント

ダイアログボックス内の3つのエラーとイベントは、 `DialogEventReceived` ホストページでイベントを発生させます。 ホストページについての通知については、「[ホストページからダイアログボックスを開く](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page)」を参照してください。

|コード番号|意味|
|:-----|:-----|
|12002|以下のいずれか:<br> - `displayDialogAsync` に渡された URL にページが存在しない。<br> -読み込みに渡されたページ。 `displayDialogAsync` ただし、ダイアログボックスは、検出または読み込みできないページにリダイレクトされたか、または無効な構文の URL に転送されています。|
|12003|ダイアログ ボックスが HTTP プロトコルを使用している URL を指していました。HTTPS が必要です。|
|12006|ダイアログボックスが閉じられました。通常は、ユーザーが [**閉じる**] ボタン**X**を選択したためです。|

コードで、呼び出し内の `DialogEventReceived` イベントのハンドラーを `displayDialogAsync` に割り当てることができます。次に簡単な例を示します。

```js
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
    function (result) {
        dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogEventReceived, processDialogEvent);
    }
);
```

各エラー コードのカスタム エラー メッセージを作成する `DialogEventReceived` イベントのハンドラーの例を、次に示します。

```js
function processDialogEvent(arg) {
    switch (arg.error) {
        case 12002:
            showNotification("The dialog box has been directed to a page that it cannot find or load, or the URL syntax is invalid.");
            break;
        case 12003:
            showNotification("The dialog box has been directed to a URL with the HTTP protocol. HTTPS is required.");            break;
        case 12006:
            showNotification("Dialog closed.");
            break;
        default:
            showNotification("Unknown error in dialog box.");
            break;
    }
}
```

この方法でエラーを処理するサンプル アドインについては、「[Office アドイン ダイアログ API の例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)」を参照してください。
