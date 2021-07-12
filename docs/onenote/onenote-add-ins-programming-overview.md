---
title: OneNote の JavaScript API のプログラミングの概要
description: Web 上の OneNote アドイン用の OneNote JavaScript API について詳しく説明します。
ms.date: 10/14/2020
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: dcfeab909bfda54a4ed5b9a748e0f8a7e7d58bb9
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/09/2021
ms.locfileid: "53348602"
---
# <a name="onenote-javascript-api-programming-overview"></a>OneNote の JavaScript API のプログラミングの概要

OneNote では、OneNote on the web アドインの JavaScript API が導入されています。OneNote オブジェクトを操作する作業ウィンドウ アドイン、コンテンツ アドイン、アドイン コマンドを作成し、Web サービスや他の Web ベースのリソースに接続できます。

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

## <a name="components-of-an-office-add-in"></a>Office アドインのコンポーネント

アドインは、2 つの基本コンポーネントで構成されます。

- Web ページと必要な任意の JavaScript、CSS、他のファイルを含む **Web アプリケーション**。これらのファイルは、Web サーバーか、Microsoft Azure などの Web ホスティング サービスでホストされます。OneNote on the web では、Web アプリケーションはブラウザー コントロールや iFrame で表示されます。

- アドインの Web ページの URL とアドインの任意のアクセス要件、設定、機能を指定する **XML マニフェスト**。このファイルは、クライアントに保存されます。OneNote アドインは、他の Office アドインと同じ [マニフェスト](../develop/add-in-manifests.md)形式を使います。

### <a name="office-add-in--manifest--webpage"></a>Office アドイン = マニフェスト + Web ページ

![Office アドインはマニフェストと Web ページによって構成されます。](../images/onenote-add-in.png)

## <a name="using-the-javascript-api"></a>JavaScript API の使用

アドインは、Office アプリケーションのランタイム コンテキストを使って、JavaScript API にアクセスします。API には次の 2 つの階層があります:

- `Application` オブジェクトを通じてアクセスされる、OneNote 固有の操作のための **アプリケーション固有の API**。
- `Document` オブジェクトを通じてアクセスされ、Office アプリケーション全体で共有される **共通 API**。

### <a name="accessing-the-application-specific-api-through-the-application-object"></a>*アプリケーション* オブジェクトを使ったアプリケーション固有の API へのアクセス

`Application` オブジェクトを使って、**ノートブック**、**セクション**、**ページ** などの OneNote オブジェクトにアクセスします。 アプリケーション固有の API を使うと、プロキシ オブジェクトでバッチ操作を実行できます。 基本的な流れは、以下のようになります。

1. コンテキストからアプリケーション インスタンスを取得します。

2. 操作する OneNote オブジェクトを表すプロキシを作成します。プロキシ オブジェクトのプロパティの読み取りや書き込みを行い、メソッドを呼び出すことにより、プロキシ オブジェクトを同期的に操作します。

3. プロキシで`load`を呼び出し、パラメーターで指定されたプロパティ値を設定します。この呼び出しは、コマンドのキューに追加されます。

   > [!NOTE]
   > API へのメソッドの呼び出し (`context.application.getActiveSection().pages;` など) も、キューに追加されます。

4. キューに置かれたすべてのコマンドをキューに置かれた順序で実行するには、`context.sync` を呼び出します。これにより、実行中のスクリプトと実際のオブジェクトの間の状態が同期されます。また、読み込まれた OneNote オブジェクトのプロパティを取得して、スクリプトで使います。追加のアクションのチェーン処理には、返された約束オブジェクトを使うことができます。

例:

```js
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

OneNote JavaScript API の `load`/`sync` パターンとその他の一般的なプラクティスの詳細については、「[アプリケーション固有の API モデルの使用](../develop/application-specific-api-model.md)」を参照してください。

[API リファレンス](../reference/overview/onenote-add-ins-javascript-reference.md)では、サポートされている OneNote オブジェクトと操作を見つけることができます。

#### <a name="onenote-javascript-api-requirement-sets"></a>OneNote JavaScript API の要件セット

要件セットは、API メンバーの名前付きグループです。 Office アドインでは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office アプリケーションがアドインに必要な API をサポートしているかどうかを判断します。 OneNote JavaScript API 要件セットの詳細については、「[OneNote JavaScript API の要件セット](../reference/requirement-sets/onenote-api-requirement-sets.md)」を参照してください。

### <a name="accessing-the-common-api-through-the-document-object"></a>*ドキュメント* オブジェクトを使った共通 API へのアクセス

`Document` オブジェクトを使って、[getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) メソッドや [setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) メソッドなどの共通 API にアクセスします。

次に例を示します。  

```js
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
| [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) | Office.CoercionType.Text`Office.CoercionType.Text` と Office.CoercionType.Matrix`Office.CoercionType.Matrix` のみ |
| [Office.context.document.setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) | `Office.CoercionType.Text`、`Office.CoercionType.Image` と `Office.CoercionType.Html` のみ | 
| 
  [var mySetting = Office.context.document.settings.get(name);](/javascript/api/office/office.settings#get-name-) | 設定はコンテンツ アドインによってのみサポートされます | 
| 
  [Office.context.document.settings.set(name, value);](/javascript/api/office/office.settings#set-name--value-) | 設定はコンテンツ アドインによってのみサポートされます | 
| [Office.EventType.DocumentSelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs) ||

一般に、アプリケーション固有の API でサポートされていない操作を行う場合は、共通 API を使用します。 共通 API の使用の詳細については、「[共通 JavaScript API オブジェクト モデル](../develop/office-javascript-api-object-model.md)」を参照してください。

<a name="om-diagram"></a>
## <a name="onenote-object-model-diagram"></a>OneNote のオブジェクト モデル図
次の図では、OneNote JavaScript API で現在使用可能なものが示されます。

  ![OneNote のオブジェクト モデル図。](../images/onenote-om.png)

## <a name="see-also"></a>関連項目

- [Office アドインを開発する](../develop/develop-overview.md)
- [Microsoft 365 開発者プログラムについて学ぶ](https://developer.microsoft.com/microsoft-365/dev-program)
- [最初の OneNote アドインをビルドする](../quickstarts/onenote-quickstart.md)
- [OneNote JavaScript API リファレンス](../reference/overview/onenote-add-ins-javascript-reference.md)
- [Rubric Grader のサンプル](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Office アドイン プラットフォームの概要](../overview/office-add-ins.md)
