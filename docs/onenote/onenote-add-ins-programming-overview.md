# <a name="onenote-javascript-api-programming-overview"></a>OneNote の JavaScript API のプログラミングの概要

OneNote では、OneNote Online アドインの JavaScript API が導入されています。OneNote オブジェクトを操作する作業ウィンドウ アドイン、コンテンツ アドイン、アドイン コマンドを作成し、Web サービスや他の Web ベースのリソースに接続できます。

>
  **注:**アドインをビルドするとき、アドインを Office ストアに[発行](../publish/publish.md)する予定であれば、[Office ストア検証ポリシー](https://msdn.microsoft.com/en-us/library/jj220035.aspx)に準拠していることを確認してください。たとえば、検証に合格するには、アドインは、定義したメソッドをサポートするすべてのプラットフォーム全体で機能する必要があります (詳細については、[セクション 4.12](https://msdn.microsoft.com/en-us/library/jj220035.aspx#Anchor_3) と「[Office アドインを使用できるホストおよびプラットフォーム](https://dev.office.com/add-in-availability)」のページを参照してください)。

## <a name="components-of-an-office-add-in"></a>Office アドインのコンポーネント

アドインは、2 つの基本コンポーネントで構成されます。

- Web ページと必要な任意の JavaScript、CSS、他のファイルで構成される **Web アプリケーション**。 これらのファイルは、Web サーバーか、Microsoft Azure などの Web ホスティング サービスでホストされます。 OneNote Online では、Web アプリケーションはブラウザー コントロールや iFrame で表示されます。
    
- アドインの Web ページの URL とアドインの任意のアクセス要件、設定、機能を指定する **XML マニフェスト**。 このファイルは、クライアントに保存されます。 OneNote アドインは、他の Office アドインと同じ[マニフェスト](https://dev.office.com/docs/add-ins/overview/add-in-manifests)形式を使います。

**Office アドイン = マニフェスト + Web ページ**

![Office アドインはマニフェストと Web ページによって構成されます](../images/onenote-add-in.png)

## <a name="using-the-javascript-api"></a>JavaScript API の使用

アドインは、ホスト アプリケーションのランタイム コンテキストを使って、JavaScript API にアクセスします。API には次の 2 つの階層があります。 

- **アプリケーション** オブジェクトを通してアクセスされる、OneNote 固有の操作のための**豊富な API**。
- **ドキュメント** オブジェクトを通してアクセスされ、Office アプリケーション全体で共有される**共通 API**。

### <a name="accessing-the-rich-api-through-the-application-object"></a>*アプリケーション* オブジェクトを使った豊富な API へのアクセス

**アプリケーション** オブジェクトを使って、**ノートブック**、**セクション**、**ページ**などの OneNote オブジェクトにアクセスします。 豊富な API を使うと、プロキシ オブジェクトでバッチ操作を実行できます。 基本的な流れは、以下のようになります。 

1. コンテキストからアプリケーション インスタンスを取得します。

2. 操作する OneNote オブジェクトを表すプロキシを作成します。プロキシ オブジェクトのプロパティの読み取りや書き込みを行い、メソッドを呼び出すことにより、プロキシ オブジェクトを同期的に操作します。 

3. プロキシで **load** を呼び出して、パラメーターで指定されたプロパティ値を設定します。 この呼び出しは、コマンドのキューに追加されます。

    > **注**:API へのメソッドの呼び出し (`context.application.getActiveSection().pages;` など) も、キューに追加されます。

4. キューに置かれたすべてのコマンドをキューに置かれた順序で実行するには、**context.sync** を呼び出します。 これにより、実行中のスクリプトと実際のオブジェクトの間の状態が同期されます。また、読み込まれた OneNote オブジェクトのプロパティを取得して、スクリプトで使います。 追加のアクションのチェーン処理には、返された約束オブジェクトを使うことができます。

例: 

```
    function getPagesInSection() {
        OneNote.run(function (context) {
            
            // Get the pages in the current section.
            var pages = context.application.getActiveSection().pages;
            
            // Queue a command to load the id and title for each page.            
            pages.load('id,title');
            
            // Run the queued commands, and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    
                    // Read the id and title of each page. 
                    $.each(pages.items, function(index, page) {
                        var pageId = page.id;
                        var pageTitle = page.title;
                        console.log(pageTitle + ': ' + pageId); 
                    });
                })
                .catch(function (error) {
                    app.showNotification("Error: " + error);
                    console.log("Error: " + error);
                    if (error instanceof OfficeExtension.Error) {
                        console.log("Debug info: " + JSON.stringify(error.debugInfo));
                    }
                });
        });
    }
```

[API リファレンス](../../reference/onenote/onenote-add-ins-javascript-reference.md)では、サポートされている OneNote オブジェクトと操作を見つけることができます。

### <a name="accessing-the-common-api-through-the-document-object"></a>*ドキュメント* オブジェクトを使った共通 API へのアクセス

**ドキュメント** オブジェクトを使って、[getSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.getselecteddataasync) メソッドや [setSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.setselecteddataasync) メソッドなどの共通 API にアクセスします。 

例:  

```
function getSelectionFromPage() {
    Office.context.document.getSelectedDataAsync(
        Office.CoercionType.Text,
        { valueFormat: "unformatted" },
        function (asyncResult) {
            var error = asyncResult.error;
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.log(error.message);
            }
            else $('#input').val(asyncResult.value);
        });
}
```
OneNote アドインは、次の共通 API のみをサポートします。

| API | メモ |
|:------|:------|
| [Office.context.document.getSelectedDataAsync](https://msdn.microsoft.com/en-us/library/office/fp142294.aspx) | **Office.CoercionType.Text** と **Office.CoercionType.Matrix** のみ |
| [Office.context.document.setSelectedDataAsync](https://msdn.microsoft.com/en-us/library/office/fp142145.aspx) | **Office.CoercionType.Text**、**Office.CoercionType.Image**、**Office.CoercionType.Html** のみ | 
| 
  [var mySetting = Office.context.document.settings.get(name);](https://msdn.microsoft.com/en-us/library/office/fp142180.aspx) | 設定はコンテンツ アドインによってのみサポートされます | 
| 
  [Office.context.document.settings.set(name, value);](https://msdn.microsoft.com/en-us/library/office/fp161063.aspx) | 設定はコンテンツ アドインによってのみサポートされます | 
| [Office.EventType.DocumentSelectionChanged](https://dev.office.com/reference/add-ins/shared/document.selectionchanged.event) ||

一般に、豊富な API でサポートされていない操作を行う場合は、共通 API のみを使います。 共通 API の使用について詳しくは、Office アドインの[ドキュメント](https://dev.office.com/docs/add-ins/overview/office-add-ins)と[リファレンス](https://dev.office.com/reference/add-ins/javascript-api-for-office)をご覧ください。


<a name="om-diagram"></a>
## <a name="onenote-object-model-diagram"></a>OneNote のオブジェクト モデル図 
次の図では、OneNote JavaScript API で現在使用可能なものが示されます。

  ![OneNote のオブジェクト モデル図](../images/onenote-om.png)


## <a name="additional-resources"></a>その他の技術情報

- [最初の OneNote 用アドインをビルドする](onenote-add-ins-getting-started.md)
- [OneNote JavaScript API リファレンス](../../reference/onenote/onenote-add-ins-javascript-reference.md)
- [Rubric Grader のサンプル](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Office アドイン プラットフォームの概要](https://dev.office.com/docs/add-ins/overview/office-add-ins)
