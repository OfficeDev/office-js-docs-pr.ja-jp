---
title: PowerPoint アドイン
description: ''
ms.date: 10/16/2018
ms.openlocfilehash: 390497e74d4dc52b9d400f242850ab72bdb0eabc
ms.sourcegitcommit: a6d6348075c1abed76d2146ddfc099b0151fe403
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/19/2018
ms.locfileid: "25640079"
---
# <a name="powerpoint-add-ins"></a>PowerPoint アドイン

PowerPoint のアドインを使って、Windows、iOS、Office Online、Mac などのプラットフォームでユーザーのプレゼンテーションのための魅力的なソリューションをビルドすることができます。 PowerPoint のアドインの 2 つの種類を作成することができます。

- **コンテンツ アドイン**を使うと、プレゼンテーションに HTML5 の動的コンテンツが追加されます。たとえば [PowerPoint のための LucidChart ダイアグラム](https://store.office.com/app.aspx?assetid=WA104380117&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Productivity&homapppos=3&homchv=2&appredirect=false) アドインでは、これを使って LucidChart からデッキにインタラクティブな図を挿入することができます。

- **作業ウィンドウ アドイン** を使えば、サービスを介して、参照情報を取り込んだり、プレゼンテーションにデータを挿入したりすることができます。 たとえば [Shutterstock イメージ 像](https://store.office.com/app.aspx?assetid=WA104380169&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Editor%2527s%2BPicks&homapppos=0&homchv=1&appredirect=false) アドインでは、これを使ってプロの写真をプレゼンテーションに追加することができます。 

## <a name="powerpoint-add-in-scenarios"></a>PowerPoint アドインのシナリオ

この記事でデモンストレーションするコード例では、PowerPoint のアドインの開発のための基本的なタスクをいくつか示します。 次の点に注意してください。

- 情報を表示するために、これらの例では Visual Studio の Office アドイン プロジェクト テンプレートに含まれる `app.showNotification` 関数を使用します。 アドインの開発に Visual Studio を使用している場合、独自のコードで`showNotification` を置き換える必要があります。 

- これらの例のうちいくつかでは、これらの関数の範囲を超えて宣言されている `Globals` も使用しています。   `var Globals = {activeViewHandler:0, firstSlideId:0};`

- これらの例を使用するには、アドインプロジェクトが [ Office.js v1.1 以降のライブラリを参照する必要がありますしている必要があります](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md)。

## <a name="detect-the-presentations-active-view-and-handle-the-activeviewchanged-event"></a>プレゼンテーションのアクティブ ビューを検出し、ActiveViewChanged イベントを処理します。

コンテンツ アドインをビルドする場合は、プレゼンテーションのアクティブ ビューを取得して、 `Office.Initialize`  ハンドラーの一部として、`ActiveViewChanged`  イベントを処理する必要があります。 

> [!NOTE]
> PowerPoint Online では、スライドショーモードが新しいセッションとして扱われるため、 [Document.ActiveViewChanged](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js) イベントは発生しません。 この場合、次のコードサンプルに示すように、アドインでロード時にアクティブ ビューをフェッチする必要があります。

コードサンプルは次の通りです。

- `getActiveFileView` 関数は、プレゼンテーションの現在のビューが "編集" ビュー (**[標準]** や **[アウトライン表示]** などの、スライドを編集できるビュー) なのか "読み取り" ビュー (** [スライド ショー] ** や ** [閲覧表示]**) なのかを返す [ Document.getActiveViewAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getactiveviewasync-options--callback-)  メソッドを呼び出します。

-  `registerActiveViewChanged` 関数は、 [addHandlerAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#addhandlerasync-eventtype--handler--options--callback-) メソッドを呼び出して[Document.ActiveViewChanged](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js) イベントのハンドラを登録します。 


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

## <a name="navigate-to-a-particular-slide-in-the-presentation"></a>プレゼンテーション内の特定のスライドに移動します。

次のコードサンプルでは、​​`getSelectedRange`関数は [Document.getSelectedDataAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getselecteddataasync-coerciontype--options--callback-) メソッドを呼び出して、`asyncResult.value` から返される JSON オブジェクトを取得します。そのオブジェクトには、**slides** という名前の配列が含まれています。  **スライド** 配列には、選択した範囲のスライド（複数のスライドが選択されていない場合は現在のスライド）の ID、タイトル、およびインデックスが含まれます。 また選択範囲内の最初のスライド ID をグローバル変数に保存します。

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

次のコードサンプルでは、`goToFirstSlide` 関数は、 [Document.goToByIdAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#gotobyidasync-id--gototype--options--callback-) メソッドを呼び出して、先で示した`getSelectedRange` 関数で識別された最初のスライドに移動します。

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

## <a name="navigate-between-slides-in-the-presentation"></a>プレゼンテーション内のスライド間を移動します。

次のコードサンプルでは、`goToSlideByIndex` 関数は**Document.goToByIdAsync** メソッドを呼び出して、プレゼンテーション内の次のスライドに移動します。

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

次のコードサンプルでは、`getFileUrl`関数は [Document.getFileProperties](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getfilepropertiesasync-options--callback-) メソッドを呼び出してプレゼンテーションファイルの URL を取得します。

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
- [PowerPoint のコード サンプル](https://developer.microsoft.com/en-us/office/gallery/?filterBy=Samples,PowerPoint)
- [コンテンツ アドインおよび作業ウィンドウ アドインでドキュメントごとにアドインの状態と設定を保存する方法](../develop/persisting-add-in-state-and-settings.md#how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins)
- [現在の選択肢にデータを読み書きする](../develop/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)
- [PowerPoint や Word 用のアドインからドキュメント全体を取得する](../powerpoint/get-the-whole-document-from-an-add-in-for-powerpoint.md)
- [PowerPoint アドインでドキュメントのテーマを使用する](use-document-themes-in-your-powerpoint-add-ins.md)
    
