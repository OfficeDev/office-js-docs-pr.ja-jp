---
title: OneNote の JavaScript API のプログラミングの概要
description: Web 上の OneNote アドイン用の OneNote JavaScript API について詳しく説明します。
ms.date: 07/18/2022
ms.topic: overview
ms.custom: scenarios:getting-started
ms.localizationpriority: high
ms.openlocfilehash: d44a01cf0f676057ca072cff74e2e80057f645f4
ms.sourcegitcommit: 05be1086deb2527c6c6ff3eafcef9d7ed90922ec
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/28/2022
ms.locfileid: "68092911"
---
# <a name="onenote-javascript-api-programming-overview"></a>OneNote の JavaScript API のプログラミングの概要

OneNote introduces a JavaScript API for OneNote add-ins on the web. You can create task pane add-ins, content add-ins, and add-in commands that interact with OneNote objects and connect to web services or other web-based resources.

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

## <a name="components-of-an-office-add-in"></a>Office アドインのコンポーネント

アドインは、2 つの基本コンポーネントで構成されます。

- A **web application** consisting of a webpage and any required JavaScript, CSS, or other files. These files are hosted on a web server or web hosting service, such as Microsoft Azure. In OneNote on the web, the web application displays in a browser control or iframe.

- An **XML manifest** that specifies the URL of the add-in's webpage and any access requirements, settings, and capabilities for the add-in. This file is stored on the client. OneNote add-ins use the same [manifest](../develop/add-in-manifests.md) format as other Office Add-ins.

### <a name="office-add-in--manifest--webpage"></a>Office アドイン = マニフェスト + Web ページ

![Office アドインはマニフェストと Web ページによって構成されます。](../images/onenote-add-in.png)

## <a name="using-the-javascript-api"></a>JavaScript API の使用

Add-ins use the runtime context of the Office application to access the JavaScript API. The API has two layers:

- `Application` オブジェクトを通じてアクセスされる、OneNote 固有の操作のための **アプリケーション固有の API**。
- `Document` オブジェクトを通じてアクセスされ、Office アプリケーション全体で共有される **共通 API**。

### <a name="accessing-the-application-specific-api-through-the-application-object"></a>*アプリケーション* オブジェクトを使ったアプリケーション固有の API へのアクセス

Use the `Application` object to access OneNote objects such as **Notebook**, **Section**, and **Page**. With application-specific APIs, you run batch operations on proxy objects. The basic flow goes something like this:

1. コンテキストからアプリケーション インスタンスを取得します。

2. Create a proxy that represents the OneNote object you want to work with. You interact synchronously with proxy objects by reading and writing their properties and calling their methods.

3. Call `load` on the proxy to fill it with the property values specified in the parameter. This call is added to the queue of commands.

   > [!NOTE]
   > API へのメソッドの呼び出し (`context.application.getActiveSection().pages;` など) も、キューに追加されます。

4. Call `context.sync` to run all queued commands in the order that they were queued. This synchronizes the state between your running script and the real objects, and by retrieving properties of loaded OneNote objects for use in your script. You can use the returned promise object for chaining additional actions.

例:

```js
async function getPagesInSection() {
    await OneNote.run(async (context) => {

        // Get the pages in the current section.
        const pages = context.application.getActiveSection().pages;

        // Queue a command to load the id and title for each page.
        pages.load('id,title');

        // Run the queued commands, and return a promise to indicate task completion.
        await context.sync();
            
        // Read the id and title of each page.
        $.each(pages.items, function(index, page) {
            let pageId = page.id;
            let pageTitle = page.title;
            console.log(pageTitle + ': ' + pageId);
        });
    });
}
```

OneNote JavaScript API の `load`/`sync` パターンとその他の一般的なプラクティスの詳細については、「[アプリケーション固有の API モデルの使用](../develop/application-specific-api-model.md)」を参照してください。

[API リファレンス](../reference/overview/onenote-add-ins-javascript-reference.md)では、サポートされている OneNote オブジェクトと操作を見つけることができます。

#### <a name="onenote-javascript-api-requirement-sets"></a>OneNote JavaScript API の要件セット

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For detailed information about OneNote JavaScript API requirement sets, see [OneNote JavaScript API requirement sets](/javascript/api/requirement-sets/onenote/onenote-api-requirement-sets).

### <a name="accessing-the-common-api-through-the-document-object"></a>*ドキュメント* オブジェクトを使った共通 API へのアクセス

`Document` オブジェクトを使って、[getSelectedDataAsync](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)) メソッドや [setSelectedDataAsync](/javascript/api/office/office.document#office-office-document-setselecteddataasync-member(1)) メソッドなどの共通 API にアクセスします。

次に例を示します。  

```js
function getSelectionFromPage() {
    Office.context.document.getSelectedDataAsync(
        Office.CoercionType.Text,
        { valueFormat: "unformatted" },
        function (asyncResult) {
            const error = asyncResult.error;
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
| [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)) | Office.CoercionType.Text`Office.CoercionType.Text` と Office.CoercionType.Matrix`Office.CoercionType.Matrix` のみ |
| [Office.context.document.setSelectedDataAsync](/javascript/api/office/office.document#office-office-document-setselecteddataasync-member(1)) | `Office.CoercionType.Text`、`Office.CoercionType.Image` と `Office.CoercionType.Html` のみ |
| [const mySetting = Office.context.document.settings.get(name);](/javascript/api/office/office.settings#office-office-settings-get-member(1)) | 設定はコンテンツ アドインによってのみサポートされます |
| 
  [Office.context.document.settings.set(name, value);](/javascript/api/office/office.settings#office-office-settings-set-member(1)) | 設定はコンテンツ アドインによってのみサポートされます |
| [Office.EventType.DocumentSelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs) |*なし。*|

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
