---
title: Office ダイアログ ボックスでのエラーとイベントの処理
description: '[エラー] ダイアログ ボックスを開いて使用するときにエラーをトラップして処理するOffice説明します。'
ms.date: 07/08/2021
localization_priority: Normal
ms.openlocfilehash: 86b8e6f3ff6dba72245d70551846884901ec597a
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/08/2021
ms.locfileid: "58936497"
---
# <a name="handle-errors-and-events-in-the-office-dialog-box"></a>[エラーとイベントの処理] ダイアログ ボックスOffice処理する

この記事では、ダイアログ ボックスを開く際にエラーをトラップして処理する方法と、ダイアログ ボックス内で発生するエラーについて説明します。

> [!NOTE]
> この記事では、「Office アドインで Office ダイアログ API を使用する」の説明に従って、Office ダイアログ[API](dialog-api-in-office-add-ins.md)の使用の基本について理解している必要があります。
> 
> 詳細については、「[ベスト プラクティスとルール」を参照Office API を参照してください](dialog-best-practices.md)。

コードは 2 つのカテゴリのイベントを処理する必要があります。

- ダイアログ ボックスを作成できないために `displayDialogAsync` の呼び出しによって返されるエラー。
- ダイアログ ボックスのエラー、その他のイベント。

## <a name="errors-from-displaydialogasync"></a>displayDialogAsync のエラー

プラットフォームとシステムの一般的なエラーに加えて、4 つのエラーは呼び出しに固有です `displayDialogAsync` 。

|コード番号|意味|
|:-----|:-----|
|12004|`displayDialogAsync` に渡される URL のドメインは信頼されていません。ドメインは、ホスト ページと同じドメインにある必要があります (プロトコルとポート番号を含む)。|
|12005|`displayDialogAsync` に渡される URL には HTTP プロトコルを使用します。 HTTPS が必要です。 (一部のバージョンの Office 12005 で返されるエラー メッセージ テキストは、12004 で返されるのと同じです)。|
|<span id="12007">12007</span><!-- The span is needed because office-js-helpers has an error message that links to this table row. -->|ダイアログ ボックスは、このホスト ウィンドウで既に開いています。作業ウィンドウなどのホスト ウィンドウで一度に開けるダイアログ ボックスは 1 つだけです。|
|12009|ダイアログ ボックスを無視するようにユーザーが選択しました。 このエラーは、ユーザー Office on the webダイアログ ボックスの表示を許可しない場合がある場合に発生する可能性があります。 詳細については、「ポップアップ ブロック[を使用したポップアップ ブロックの処理」を参照Office on the web。](dialog-best-practices.md#handle-pop-up-blockers-with-office-on-the-web)|

呼 `displayDialogAsync` び出された場合 [、AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトをコールバック関数に渡します。 呼び出しが成功すると、ダイアログ ボックスが開き、オブジェクトのプロパティ `value` `AsyncResult` が [Dialog](/javascript/api/office/office.dialog) オブジェクトになります。 この例については、「ダイアログ ボックスから [ホスト ページに情報を送信する」を参照してください](dialog-api-in-office-add-ins.md#send-information-from-the-dialog-box-to-the-host-page)。 呼び出しが失敗すると、ダイアログ ボックスは作成されません。オブジェクトのプロパティはに設定され、オブジェクトの `displayDialogAsync` `status` `AsyncResult` `Office.AsyncResultStatus.Failed` `error` プロパティが設定されます。 エラーが発生した場合は、常にテストし `status` 、応答するコールバックを指定する必要があります。 コード番号に関係なくエラー メッセージを報告する例については、次のコードを参照してください。 (この `showNotification` 記事で定義されていない関数は、エラーを表示またはログに記録します。 アドイン内でこの関数を実装する方法の例については、「Office ダイアログ API の例 」[を参照してください](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)。

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

## <a name="errors-and-events-in-the-dialog-box"></a>ダイアログ ボックスのエラーとイベント

ダイアログ ボックス内の 3 つのエラーとイベントは、ホスト `DialogEventReceived` ページでイベントを発生します。 ホスト ページの種類を確認するには、「ホスト ページからダイアログ ボックスを開 [く」を参照してください](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page)。

|コード番号|意味|
|:-----|:-----|
|12002|以下のいずれか:<br> - `displayDialogAsync` に渡された URL にページが存在しない。<br> - 読み込まれたが、ダイアログ ボックスが見つからまたは読み込めないページにリダイレクトされたページ、または構文が無効な URL にリダイレクト `displayDialogAsync` されたページ。|
|12003|ダイアログ ボックスが HTTP プロトコルを使用している URL を指していました。HTTPS が必要です。|
|12006|ダイアログ ボックスが閉じられました。通常、ユーザーが [閉じる] ボタン **X を選択****したためです**。|

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

エラー コードごとにカスタム エラー メッセージを作成するイベントのハンドラーの例については、 `DialogEventReceived` 次の例を参照してください。

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
