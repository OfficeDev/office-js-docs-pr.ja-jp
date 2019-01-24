---
title: PowerPoint アドイン
description: ''
ms.date: 10/16/2018
localization_priority: Priority
ms.openlocfilehash: 022bed349dde061b61a8db0711a94a0a4d77f2e1
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/23/2019
ms.locfileid: "29388634"
---
# <a name="powerpoint-add-ins"></a>PowerPoint アドイン

PowerPoint のアドインを使って、Windows、iOS、Office Online、Mac などのプラットフォームでユーザーのプレゼンテーションのための魅力的なソリューションを構築することができます。 次の 2 種類の PowerPoint アドインを作成できます。

- **コンテンツ アドイン**を使うと、プレゼンテーションに HTML5 の動的コンテンツが追加されます。たとえば [PowerPoint のための LucidChart ダイアグラム](https://store.office.com/app.aspx?assetid=WA104380117&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Productivity&homapppos=3&homchv=2&appredirect=false) アドインでは、これを使って LucidChart からデッキにインタラクティブな図を挿入することができます。

- **作業ウィンドウ アドイン**を使えば、サービスを介して、参照情報を取り込んだり、プレゼンテーションにデータを挿入したりすることができます。 たとえば [Shutterstock イメージ](https://store.office.com/app.aspx?assetid=WA104380169&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Editor%2527s%2BPicks&homapppos=0&homchv=1&appredirect=false) アドインでは、これを使ってプロの写真をプレゼンテーションに追加することができます。 

## <a name="powerpoint-add-in-scenarios"></a>PowerPoint アドインのシナリオ

この記事で紹介するコード例では、PowerPoint のアドインの開発のための基本的なタスクをいくつか示します。 以下のことに注意してください。

- 情報を表示するために、これらの例は `app.showNotification` 関数を使用します。これは、Visual Studio の Office アドイン プロジェクト テンプレートに含まれています。 アドインを開発するのに Visual Studio を使用していない場合は、`showNotification` 関数を独自のコードに置き換える必要があります。 

- これらの例のうちいくつかでは、これらの関数 `var Globals = {activeViewHandler:0, firstSlideId:0};` の範囲を超えて宣言されている `Globals` オブジェクトも使用しています。

- これらの例を使用するには、アドイン プロジェクトで [Office.js v1.1 ライブラリ以降を参照](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md)する必要があります。

## <a name="detect-the-presentations-active-view-and-handle-the-activeviewchanged-event"></a>プレゼンテーションのアクティブ ビューの検出と ActiveViewChanged イベントの処理を行う

コンテンツ アドインをビルドする場合は、プレゼンテーションのアクティブ ビューを取得して、`Office.Initialize` ハンドラーの一部として、`ActiveViewChanged` イベントを処理する必要があります。 

> [!NOTE]
> PowerPoint Online では [Document.ActiveViewChanged](https://docs.microsoft.com/javascript/api/office/office.document) イベントは、スライド ショー モードが新しいセッションとして扱われるようには起動しません。 この場合、次のコード サンプルに示すように、アドインで読み込むアクティブ ビューをフェッチする必要があります。

コード サンプルは次のとおりです。

- `getActiveFileView` 関数は [Document.getActiveViewAsync](https://docs.microsoft.com/javascript/api/office/office.document#getactiveviewasync-options--callback-) メソッドを呼び出して、プレゼンテーションの現在のビューが "編集" ビュー (**[標準]** や **[アウトライン表示]** などの、スライドを編集できるビュー) なのか "読み取り" ビュー (**[スライド ショー]** や **[閲覧表示]**) なのかを返します。

- `registerActiveViewChanged` 関数は、[Document.ActiveViewChanged](https://docs.microsoft.com/javascript/api/office/office.document) イベントのハンドラーを登録するための [addHandlerAsync](https://docs.microsoft.com/javascript/api/office/office.document#addhandlerasync-eventtype--handler--options--callback-) メソッドを呼び出します。 


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

次のコード サンプルでは、`getSelectedRange` 関数は [Document.getSelectedDataAsync](https://docs.microsoft.com/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) メソッドを呼び出して、`asyncResult.value` によって返される JSON オブジェクトを取得します。このオブジェクトには、**slides** という名前の配列が含まれます。 **slides** 配列には、選択した範囲のスライド (複数のスライドが選択されていない場合は現在のスライド) の ID、タイトル、およびインデックスが含まれます。 また、選択範囲内の最初のスライド ID をグローバル変数に保存します。

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

次のコード サンプルでは、`goToFirstSlide` 関数は [Document.goToByIdAsync](https://docs.microsoft.com/javascript/api/office/office.document#gotobyidasync-id--gototype--options--callback-) メソッドを呼び出して、前に示した `getSelectedRange` 関数で識別された最初のスライドに移動します。

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

次のコード サンプルでは、`goToSlideByIndex` 関数は **Document.goToByIdAsync** メソッドを呼び出して、プレゼンテーションの次のスライドに移動します。

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

次のコード サンプルでは、`getFileUrl` 関数は [Document.getFileProperties](https://docs.microsoft.com/javascript/api/office/office.document#getfilepropertiesasync-options--callback-) メソッドを呼び出して、プレゼンテーション ファイルの URL を取得します。

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



## <a name="see-also"></a>関連項目
- 
  [PowerPoint のコード サンプル](https://developer.microsoft.com/en-us/office/gallery/?filterBy=Samples,PowerPoint)
- [コンテンツ アドインおよび作業ウィンドウ アドインで、ドキュメントごとにアドインの状態と設定を保存する方法](../develop/persisting-add-in-state-and-settings.md#how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins)
- [ドキュメントやスプレッドシート内のアクティブな選択範囲へのデータの読み取りおよび書き込み](../develop/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)
- [PowerPoint や Word 用のアドインからドキュメント全体を取得する](../powerpoint/get-the-whole-document-from-an-add-in-for-powerpoint.md)
- [PowerPoint アドインでドキュメントのテーマを使用する](use-document-themes-in-your-powerpoint-add-ins.md)
    
