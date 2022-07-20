---
title: Office ダイアログ ボックスでのエラーとイベントの処理
description: Office ダイアログ ボックスを開いて使用するときに、エラーをトラップして処理する方法について説明します。
ms.date: 07/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: 0e8eefe4ee868a3cdc52ee8d425271435404bc04
ms.sourcegitcommit: df7964b6509ee6a807d754fbe895d160bc52c2d3
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/20/2022
ms.locfileid: "66889458"
---
# <a name="handle-errors-and-events-in-the-office-dialog-box"></a>Office ダイアログ ボックスでエラーとイベントを処理する

この記事では、ダイアログ ボックスを開くときにエラーをトラップして処理する方法と、ダイアログ ボックス内で発生するエラーについて説明します。

> [!NOTE]
> この記事では、「Office アドインで Office ダイアログ API を使用する」の説明に従って、 [Office ダイアログ API の使用](dialog-api-in-office-add-ins.md)の基本について理解していることを前提にしています。
>
> [Office ダイアログ API のベスト プラクティスとルール](dialog-best-practices.md)も参照してください。

コードでは、2 つのカテゴリのイベントを処理する必要があります。

- ダイアログ ボックスを作成できないために `displayDialogAsync` の呼び出しによって返されるエラー。
- ダイアログ ボックスのエラーとその他のイベント。

## <a name="errors-from-displaydialogasync"></a>displayDialogAsync のエラー

一般的なプラットフォームエラーとシステム エラーに加えて、4 つのエラーは呼び出し `displayDialogAsync`に固有です。

|コード番号|意味|
|:-----|:-----|
|12004|`displayDialogAsync` に渡される URL のドメインは信頼されていません。ドメインは、ホスト ページと同じドメインにある必要があります (プロトコルとポート番号を含む)。|
|12005|`displayDialogAsync` に渡される URL には HTTP プロトコルを使用します。 HTTPS が必要です。 (一部のバージョンの Office では、12005 で返されるエラー メッセージ テキストは、12004 で返されたものと同じです)。|
|<span id="12007">12007</span><!-- The span is needed because office-js-helpers has an error message that links to this table row. -->|ダイアログ ボックスは、このホスト ウィンドウで既に開いています。作業ウィンドウなどのホスト ウィンドウで一度に開けるダイアログ ボックスは 1 つだけです。|
|12009|ダイアログ ボックスを無視するようにユーザーが選択しました。 このエラーはOffice on the webで発生する可能性があります。ユーザーはアドインにダイアログ ボックスの表示を許可しないことを選択できます。 詳細については、「[Office on the webを使用したポップアップ ブロックの処理](dialog-best-practices.md#handle-pop-up-blockers-with-office-on-the-web)」を参照してください。|

呼び出されると `displayDialogAsync` 、 [AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトがコールバック関数に渡されます。 呼び出しが成功すると、ダイアログ ボックスが開き `value` 、オブジェクトの `AsyncResult` プロパティは [Dialog](/javascript/api/office/office.dialog) オブジェクトです。 この例については、「 [ダイアログ ボックスからホスト ページに情報を送信する」](dialog-api-in-office-add-ins.md#send-information-from-the-dialog-box-to-the-host-page)を参照してください。 呼び出し `displayDialogAsync` が失敗すると、ダイアログ ボックスは作成されず、 `status` オブジェクトの `AsyncResult` プロパティが設定 `Office.AsyncResultStatus.Failed`され `error` 、オブジェクトのプロパティが設定されます。 エラーが発生した場合にテストして応答するコールバックを `status` 常に指定する必要があります。 コード番号に関係なくエラー メッセージを報告する例については、次のコードを参照してください。 (この記事では定義されていない関数は `showNotification` 、エラーを表示またはログに記録します。 アドイン内でこの関数を実装する方法の例については、「 [Office アドイン ダイアログ API の例」を参照してください](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example))。

```js
let dialog;
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

## <a name="errors-and-events-in-the-dialog-box"></a>ダイアログ ボックスのエラーとイベント

ダイアログ ボックスの 3 つのエラーとイベントによって、ホスト ページにイベントが発生 `DialogEventReceived` します。 ホスト ページの概要については、「ホスト ページ [からダイアログ ボックスを開く](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page)」を参照してください。

|コード番号|意味|
|:-----|:-----|
|12002|以下のいずれか:<br> - `displayDialogAsync` に渡された URL にページが存在しない。<br> - 読み込まれたが `displayDialogAsync` 、ダイアログ ボックスが見つからないか読み込めないページにリダイレクトされたか、構文が無効な URL にリダイレクトされたページ。|
|12003|ダイアログ ボックスが HTTP プロトコルを使用している URL を指していました。HTTPS が必要です。|
|12006|通常、ユーザーが **[閉じる** ] ボタン **[X**] を選択したため、ダイアログ ボックスは閉じられました。|

コードで、呼び出し内の `DialogEventReceived` イベントのハンドラーを `displayDialogAsync` に割り当てることができます。 次に簡単な例を示します。

```js
let dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
    function (result) {
        dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogEventReceived, processDialogEvent);
    }
);
```

エラー コードごとにカスタム エラー メッセージを `DialogEventReceived` 作成するイベントのハンドラーの例については、次の例を参照してください。

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

## <a name="see-also"></a>関連項目

この方法でエラーを処理するサンプル アドインについては、「[Office アドイン ダイアログ API の例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)」を参照してください。
