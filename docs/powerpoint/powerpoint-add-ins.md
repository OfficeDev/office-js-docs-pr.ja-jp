---
title: PowerPoint アドイン
description: PowerPoint アドインを使用して Windows、iPad、Mac、ブラウザー上など、複数のプラットフォームでプレゼンテーション用の魅力的なソリューションを構築する方法を説明します。
ms.date: 10/14/2020
ms.topic: overview
ms.custom: scenarios:getting-started
ms.localizationpriority: high
ms.openlocfilehash: a1bfa8f17f6a63896025a374a9fe8a6bdbf36f55
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/23/2022
ms.locfileid: "63746995"
---
# <a name="powerpoint-add-ins"></a>PowerPoint アドイン

PowerPoint アドインを使用して、Windows、iPad、Mac を含むプラットフォーム間、およびブラウザーで、ユーザーのプレゼンテーション用に魅力的なソリューションを構築できます。次の 2 種類の PowerPoint アドインを作成できます。

- **コンテンツ アドイン** を使うと、プレゼンテーションに HTML5 の動的コンテンツが追加されます。たとえば [PowerPoint のための LucidChart ダイアグラム](https://appsource.microsoft.com/product/office/wa104380117) アドインでは、これを使って LucidChart からデッキにインタラクティブな図を挿入することができます。

- **作業ウィンドウ アドイン** を使えば、サービスを介して、参照情報を取り込んだり、スライドにデータを挿入したりすることができます。たとえば [Pexels - Free Stock Photos](https://appsource.microsoft.com/product/office/wa104379997) アドインでは、これを使ってプロの写真をプレゼンテーションに追加することができます。

## <a name="powerpoint-add-in-scenarios"></a>PowerPoint アドインのシナリオ

この記事で紹介するコード例では、PowerPoint のアドインの開発のための基本的なタスクをいくつか示します。次の点に注意してください。

- 情報を表示するために、これらの例は `app.showNotification` 関数を使用しています。これは、Visual Studio の Office アドイン プロジェクト テンプレートに含まれています。アドインを開発するのに Visual Studio を使用していない場合は、`showNotification` 関数を独自のコードに置き換える必要があります。

- これらの例のうちいくつかでは、これらの関数 `var Globals = {activeViewHandler:0, firstSlideId:0};` の範囲を超えて宣言されている `Globals` オブジェクトも使用しています。

- これらの例を使用するには、アドイン プロジェクトで [Office.js v1.1 ライブラリ以降を参照](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md)する必要があります。

## <a name="detect-the-presentations-active-view-and-handle-the-activeviewchanged-event"></a>プレゼンテーションのアクティブ ビューの検出と ActiveViewChanged イベントの処理を行う

コンテンツ アドインをビルドする場合は、プレゼンテーションのアクティブ ビューを取得して、`Office.Initialize` ハンドラーの一部として、`ActiveViewChanged` イベントを処理する必要があります。

> [!NOTE]
> PowerPoint on the web では [Document.ActiveViewChanged](/javascript/api/office/office.document) イベントは、スライド ショー モードが新しいセッションとして扱われるので、起動しません。この場合、下のコードサンプルで示すように、アドインでアクティブ ビューを読み込むようにすることが必要です。

コード サンプルは次のとおりです。

- `getActiveFileView` 関数は [Document.getActiveViewAsync](/javascript/api/office/office.document#office-office-document-getactiveviewasync-member(1)) メソッドを呼び出して、プレゼンテーションの現在のビューが "編集" ビュー (**[標準]** や **[アウトライン表示]** などの、スライドを編集できるビュー) なのか "読み取り" ビュー (**[スライド ショー]** や **[閲覧表示]**) なのかを返します。

- `registerActiveViewChanged` 関数は、[Document.ActiveViewChanged](/javascript/api/office/office.document) イベントのハンドラーを登録するための [addHandlerAsync](/javascript/api/office/office.document#office-office-document-addhandlerasync-member(1)) メソッドを呼び出します。


```js
//general Office.initialize function. Fires on load of the add-in.
Office.initialize = function(){

    //Gets whether the current view is edit or read.
    var currentView = getActiveFileView();

    //register for the active view changed handler
    registerActiveViewChanged();

    //render the content based off of the currentView
    //....
}

function getActiveFileView()
{
    Office.context.document.getActiveViewAsync(function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            app.showNotification(asyncResult.value);
        }
    });

}

function registerActiveViewChanged() {
    Globals.activeViewHandler = function (args) {
        app.showNotification(JSON.stringify(args));
    }

    Office.context.document.addHandlerAsync(Office.EventType.ActiveViewChanged, Globals.activeViewHandler,
        function (asyncResult) {
            if (asyncResult.status == "failed") {
                app.showNotification("Action failed with error: " + asyncResult.error.message);
            }
            else {
                app.showNotification(asyncResult.status);
            }
        });
}
```

## <a name="navigate-to-a-particular-slide-in-the-presentation"></a>プレゼンテーションの特定のスライドに移動する

次のコードサンプルでは、`getSelectedRange`関数は [Document.getSelectedDataAsync](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)) メソッドを呼び出して、`asyncResult.value`から返される JSON オブジェクトを取得します。そのオブジェクトには、`slides`という名前の配列が含まれています。`slides` の配列には、選択範囲のスライド (または複数のスライドが選択されていない場合は、現在のスライドのみ) の ID、タイトル、インデックスが含まれていいます。この関数はまた、選択範囲の最初のスライドの ID をグローバル変数に保存します。

```js
function getSelectedRange() {
    // Get the id, title, and index of the current slide (or selected slides) and store the first slide id */
    Globals.firstSlideId = 0;

    Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            Globals.firstSlideId = asyncResult.value.slides[0].id;
            app.showNotification(JSON.stringify(asyncResult.value));
        }
    });
}
```

次のコード サンプルでは、`goToFirstSlide` 関数は [Document.goToByIdAsync](/javascript/api/office/office.document#office-office-document-gotobyidasync-member(1)) メソッドを呼び出して、前に示した `getSelectedRange` 関数で識別された最初のスライドに移動します。

```js
function goToFirstSlide() {
    Office.context.document.goToByIdAsync(Globals.firstSlideId, Office.GoToType.Slide, function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            app.showNotification("Navigation successful");
        }
    });
}
```

## <a name="navigate-between-slides-in-the-presentation"></a>プレゼンテーション内のスライド間を移動する

次のコード サンプルでは、`goToSlideByIndex` 関数は `Document.goToByIdAsync` メソッドを呼び出して、プレゼンテーションの次のスライドに移動します。

```js
function goToSlideByIndex() {
    var goToFirst = Office.Index.First;
    var goToLast = Office.Index.Last;
    var goToPrevious = Office.Index.Previous;
    var goToNext = Office.Index.Next;

    Office.context.document.goToByIdAsync(goToNext, Office.GoToType.Index, function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            app.showNotification("Navigation successful");
        }
    });
}
```

## <a name="get-the-url-of-the-presentation"></a>プレゼンテーションの URL を取得する

次のコード サンプルでは、`getFileUrl` 関数は [Document.getFileProperties](/javascript/api/office/office.document#office-office-document-getfilepropertiesasync-member(1)) メソッドを呼び出して、プレゼンテーション ファイルの URL を取得します。

```js
function getFileUrl() {
    //Get the URL of the current file.
    Office.context.document.getFilePropertiesAsync(function (asyncResult) {
        var fileUrl = asyncResult.value.url;
        if (fileUrl == "") {
            app.showNotification("The file hasn't been saved yet. Save the file and try again");
        }
        else {
            app.showNotification(fileUrl);
        }
    });
}
```

## <a name="create-a-presentation"></a>プレゼンテーションの作成

アドインでは、アドインが現在実行されている PowerPoint のインスタンスとは異なる新しいプレゼンテーションを作成できます。 PowerPoint の名前空間には、この目的のための `createPresentation` メソッドがあります。 このメソッドが呼び出されると、新しいプレゼンテーションが PowerPoint の新しいインスタンスですぐに開いて表示されます。 アドインは前のプレゼンテーションで開いて実行されたままになります。

```js
PowerPoint.createPresentation();
```

`createPresentation` メソッドでは既存のプレゼンテーションのコピーの作成もできます。 このメソッドは、オプションのパラメーターとして .pptx ファイルの base64 エンコード文字列表現を受け取ります。 文字列の引数は有効な .pptx ファイルと見なされ、作成されるプレゼンテーションはそのファイルのコピーになります。 次の例に示すように、[FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) クラスを使用して、ファイルを必要な base64 エンコード文字列に変換できます。

```js
var myFile = document.getElementById("file");
var reader = new FileReader();

reader.onload = function (event) {
    // strip off the metadata before the base64-encoded string
    var startIndex = reader.result.toString().indexOf("base64,");
    var copyBase64 = reader.result.toString().substr(startIndex + 7);

    PowerPoint.createPresentation(copyBase64);
};

// read in the file as a data URL so we can parse the base64-encoded string
reader.readAsDataURL(myFile.files[0]);
```

## <a name="see-also"></a>関連項目

- [Office アドインを開発する](../develop/develop-overview.md)
- [Microsoft 365 開発者プログラムについて](https://developer.microsoft.com/microsoft-365/dev-program)
- [PowerPoint のコード サンプル](https://developer.microsoft.com/office/gallery/?filterBy=Samples,PowerPoint)
- [コンテンツ アドインおよび作業ウィンドウ アドインで、ドキュメントごとにアドインの状態と設定を保存する方法](../develop/persisting-add-in-state-and-settings.md#how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins)
- [ドキュメントやスプレッドシート内のアクティブな選択範囲へのデータの読み取りおよび書き込み](../develop/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)
- [PowerPoint や Word 用のアドインからドキュメント全体を取得する](../powerpoint/get-the-whole-document-from-an-add-in-for-powerpoint.md)
- [PowerPoint アドインでドキュメントのテーマを使用する](use-document-themes-in-your-powerpoint-add-ins.md)
