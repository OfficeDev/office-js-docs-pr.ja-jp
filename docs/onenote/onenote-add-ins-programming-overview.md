---
title: OneNote の JavaScript API のプログラミングの概要
description: Web 上の OneNote アドイン用の OneNote JavaScript API について詳しく説明します。
ms.date: 07/28/2020
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: 86a388c2e8a89b97d82083964e0fcf120f310e66
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/28/2020
ms.locfileid: "47294060"
---
# <a name="onenote-javascript-api-programming-overview"></a><span data-ttu-id="b4dff-103">OneNote の JavaScript API のプログラミングの概要</span><span class="sxs-lookup"><span data-stu-id="b4dff-103">OneNote JavaScript API programming overview</span></span>

<span data-ttu-id="b4dff-104">OneNote では、OneNote on the web アドインの JavaScript API が導入されています。</span><span class="sxs-lookup"><span data-stu-id="b4dff-104">OneNote introduces a JavaScript API for OneNote add-ins on the web.</span></span> <span data-ttu-id="b4dff-105">OneNote オブジェクトを操作する作業ウィンドウ アドイン、コンテンツ アドイン、アドイン コマンドを作成し、Web サービスやその他の Web ベースのリソースに接続できます。</span><span class="sxs-lookup"><span data-stu-id="b4dff-105">You can create task pane add-ins, content add-ins, and add-in commands that interact with OneNote objects and connect to web services or other web-based resources.</span></span>

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

## <a name="components-of-an-office-add-in"></a><span data-ttu-id="b4dff-106">Office アドインのコンポーネント</span><span class="sxs-lookup"><span data-stu-id="b4dff-106">Components of an Office Add-in</span></span>

<span data-ttu-id="b4dff-107">アドインは、2 つの基本コンポーネントで構成されます。</span><span class="sxs-lookup"><span data-stu-id="b4dff-107">Add-ins consist of two basic components:</span></span>

- <span data-ttu-id="b4dff-108">Web ページと必要な任意の JavaScript、CSS、他のファイルで構成される **Web アプリケーション**。</span><span class="sxs-lookup"><span data-stu-id="b4dff-108">A **web application** consisting of a webpage and any required JavaScript, CSS, or other files.</span></span> <span data-ttu-id="b4dff-109">これらのファイルは、Web サーバーか、Microsoft Azure などの Web ホスティング サービスでホストされます。</span><span class="sxs-lookup"><span data-stu-id="b4dff-109">These files are hosted on a web server or web hosting service, such as Microsoft Azure.</span></span> <span data-ttu-id="b4dff-110">OneNote on the web では、Web アプリケーションはブラウザー コントロールや iFrame で表示されます。</span><span class="sxs-lookup"><span data-stu-id="b4dff-110">In OneNote on the web, the web application displays in a browser control or iframe.</span></span>

- <span data-ttu-id="b4dff-p103">アドインの Web ページの URL とアドインの任意のアクセス要件、設定、機能を指定する **XML マニフェスト**。このファイルは、クライアントに保存されます。OneNote アドインは、他の Office アドインと同じ[マニフェスト](../develop/add-in-manifests.md)形式を使います。</span><span class="sxs-lookup"><span data-stu-id="b4dff-p103">An **XML manifest** that specifies the URL of the add-in's webpage and any access requirements, settings, and capabilities for the add-in. This file is stored on the client. OneNote add-ins use the same [manifest](../develop/add-in-manifests.md) format as other Office Add-ins.</span></span>

### <a name="office-add-in--manifest--webpage"></a><span data-ttu-id="b4dff-114">Office アドイン = マニフェスト + Web ページ</span><span class="sxs-lookup"><span data-stu-id="b4dff-114">Office Add-in = Manifest + Webpage</span></span>

![Office アドインはマニフェストと Web ページによって構成されます](../images/onenote-add-in.png)

## <a name="using-the-javascript-api"></a><span data-ttu-id="b4dff-116">JavaScript API の使用</span><span class="sxs-lookup"><span data-stu-id="b4dff-116">Using the JavaScript API</span></span>

<span data-ttu-id="b4dff-p104">アドインは、Office アプリケーションのランタイム コンテキストを使って、JavaScript API にアクセスします。API には次の 2 つの階層があります:</span><span class="sxs-lookup"><span data-stu-id="b4dff-p104">Add-ins use the runtime context of the Office application to access the JavaScript API. The API has two layers:</span></span>

- <span data-ttu-id="b4dff-119">`Application` オブジェクトを通じてアクセスされる、OneNote 固有の操作のための **アプリケーション固有の API**。</span><span class="sxs-lookup"><span data-stu-id="b4dff-119">A **application-specific API** for OneNote-specific operations, accessed through the `Application` object.</span></span>
- <span data-ttu-id="b4dff-120">`Document` オブジェクトを通じてアクセスされ、Office アプリケーション全体で共有される**共通 API**。</span><span class="sxs-lookup"><span data-stu-id="b4dff-120">A **Common API** that's shared across Office applications, accessed through the `Document` object.</span></span>

### <a name="accessing-the-application-specific-api-through-the-application-object"></a><span data-ttu-id="b4dff-121">*アプリケーション* オブジェクトを使ったアプリケーション固有の API へのアクセス</span><span class="sxs-lookup"><span data-stu-id="b4dff-121">Accessing the application-specific API through the *Application* object</span></span>

<span data-ttu-id="b4dff-122">`Application` オブジェクトを使って、**ノートブック**、**セクション**、**ページ**などの OneNote オブジェクトにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="b4dff-122">Use the `Application` object to access OneNote objects such as **Notebook**, **Section**, and **Page**.</span></span> <span data-ttu-id="b4dff-123">アプリケーション固有の API を使うと、プロキシ オブジェクトでバッチ操作を実行できます。</span><span class="sxs-lookup"><span data-stu-id="b4dff-123">With application-specific APIs, you run batch operations on proxy objects.</span></span> <span data-ttu-id="b4dff-124">基本的な流れは、以下のようになります。</span><span class="sxs-lookup"><span data-stu-id="b4dff-124">The basic flow goes something like this:</span></span>

1. <span data-ttu-id="b4dff-125">コンテキストからアプリケーション インスタンスを取得します。</span><span class="sxs-lookup"><span data-stu-id="b4dff-125">Get the application instance from the context.</span></span>

2. <span data-ttu-id="b4dff-p106">操作する OneNote オブジェクトを表すプロキシを作成します。プロキシ オブジェクトのプロパティの読み取りや書き込みを行い、メソッドを呼び出すことにより、プロキシ オブジェクトを同期的に操作します。</span><span class="sxs-lookup"><span data-stu-id="b4dff-p106">Create a proxy that represents the OneNote object you want to work with. You interact synchronously with proxy objects by reading and writing their properties and calling their methods.</span></span>

3. <span data-ttu-id="b4dff-p107">プロキシで`load`を呼び出し、パラメーターで指定されたプロパティ値を設定します。この呼び出しは、コマンドのキューに追加されます。</span><span class="sxs-lookup"><span data-stu-id="b4dff-p107">Call `load` on the proxy to fill it with the property values specified in the parameter. This call is added to the queue of commands.</span></span>

   > [!NOTE]
   > <span data-ttu-id="b4dff-130">API へのメソッドの呼び出し (`context.application.getActiveSection().pages;` など) も、キューに追加されます。</span><span class="sxs-lookup"><span data-stu-id="b4dff-130">Method calls to the API (such as `context.application.getActiveSection().pages;`) are also added to the queue.</span></span>

4. <span data-ttu-id="b4dff-p108">キューに置かれたすべてのコマンドをキューに置かれた順序で実行するには、`context.sync` を呼び出します。これにより、実行中のスクリプトと実際のオブジェクトの間の状態が同期されます。また、読み込まれた OneNote オブジェクトのプロパティを取得して、スクリプトで使います。追加のアクションのチェーン処理には、返された約束オブジェクトを使うことができます。</span><span class="sxs-lookup"><span data-stu-id="b4dff-p108">Call `context.sync` to run all queued commands in the order that they were queued. This synchronizes the state between your running script and the real objects, and by retrieving properties of loaded OneNote objects for use in your script. You can use the returned promise object for chaining additional actions.</span></span>

<span data-ttu-id="b4dff-134">例:</span><span class="sxs-lookup"><span data-stu-id="b4dff-134">For example:</span></span>

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

<span data-ttu-id="b4dff-135">OneNote JavaScript API の `load`/`sync` パターンとその他の一般的なプラクティスの詳細については、「[アプリケーション固有の API モデルの使用](../develop/application-specific-api-model.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="b4dff-135">See [Using the application-specific API model](../develop/application-specific-api-model.md) to learn more about the `load`/`sync` pattern and other common practices in the OneNote JavaScript APIs.</span></span>

<span data-ttu-id="b4dff-136">[API リファレンス](../reference/overview/onenote-add-ins-javascript-reference.md)では、サポートされている OneNote オブジェクトと操作を見つけることができます。</span><span class="sxs-lookup"><span data-stu-id="b4dff-136">You can find supported OneNote objects and operations in the [API reference](../reference/overview/onenote-add-ins-javascript-reference.md).</span></span>

#### <a name="onenote-javascript-api-requirement-sets"></a><span data-ttu-id="b4dff-137">OneNote JavaScript API の要件セット</span><span class="sxs-lookup"><span data-stu-id="b4dff-137">OneNote JavaScript API requirement sets</span></span>

<span data-ttu-id="b4dff-138">要件セットは、API メンバーの名前付きグループです。</span><span class="sxs-lookup"><span data-stu-id="b4dff-138">Requirement sets are named groups of API members.</span></span> <span data-ttu-id="b4dff-139">Office アドインでは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office アプリケーションがアドインに必要な API をサポートしているかどうかを判断します。</span><span class="sxs-lookup"><span data-stu-id="b4dff-139">Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs.</span></span> <span data-ttu-id="b4dff-140">OneNote JavaScript API 要件セットの詳細については、「[OneNote JavaScript API の要件セット](../reference/requirement-sets/onenote-api-requirement-sets.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="b4dff-140">For detailed information about OneNote JavaScript API requirement sets, see [OneNote JavaScript API requirement sets](../reference/requirement-sets/onenote-api-requirement-sets.md).</span></span>

### <a name="accessing-the-common-api-through-the-document-object"></a><span data-ttu-id="b4dff-141">*ドキュメント* オブジェクトを使った共通 API へのアクセス</span><span class="sxs-lookup"><span data-stu-id="b4dff-141">Accessing the Common API through the *Document* object</span></span>

<span data-ttu-id="b4dff-142">`Document` オブジェクトを使って、[getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) メソッドや [setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) メソッドなどの共通 API にアクセスします。</span><span class="sxs-lookup"><span data-stu-id="b4dff-142">Use the `Document` object to access the Common API, such as the [getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) and [setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) methods.</span></span>

<span data-ttu-id="b4dff-143">次に例を示します。</span><span class="sxs-lookup"><span data-stu-id="b4dff-143">For example:</span></span>  

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

<span data-ttu-id="b4dff-144">OneNote アドインは、次の共通 API のみをサポートします。</span><span class="sxs-lookup"><span data-stu-id="b4dff-144">OneNote add-ins support only the following Common APIs:</span></span>

| <span data-ttu-id="b4dff-145">API</span><span class="sxs-lookup"><span data-stu-id="b4dff-145">API</span></span> | <span data-ttu-id="b4dff-146">メモ</span><span class="sxs-lookup"><span data-stu-id="b4dff-146">Notes</span></span> |
|:------|:------|
| [<span data-ttu-id="b4dff-147">Office.context.document.getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="b4dff-147">Office.context.document.getSelectedDataAsync</span></span>](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) | <span data-ttu-id="b4dff-148">Office.CoercionType.Text`Office.CoercionType.Text` と Office.CoercionType.Matrix`Office.CoercionType.Matrix` のみ</span><span class="sxs-lookup"><span data-stu-id="b4dff-148">`Office.CoercionType.Text` and `Office.CoercionType.Matrix` only</span></span> |
| [<span data-ttu-id="b4dff-149">Office.context.document.setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="b4dff-149">Office.context.document.setSelectedDataAsync</span></span>](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) | <span data-ttu-id="b4dff-150">`Office.CoercionType.Text`、`Office.CoercionType.Image` と `Office.CoercionType.Html` のみ</span><span class="sxs-lookup"><span data-stu-id="b4dff-150">`Office.CoercionType.Text`, `Office.CoercionType.Image`, and `Office.CoercionType.Html` only</span></span> | 
| <span data-ttu-id="b4dff-151">
  [var mySetting = Office.context.document.settings.get(name);](/javascript/api/office/office.settings#get-name-)</span><span class="sxs-lookup"><span data-stu-id="b4dff-151">[var mySetting = Office.context.document.settings.get(name);](/javascript/api/office/office.settings#get-name-)</span></span> | <span data-ttu-id="b4dff-152">設定はコンテンツ アドインによってのみサポートされます</span><span class="sxs-lookup"><span data-stu-id="b4dff-152">Settings are supported by content add-ins only</span></span> | 
| <span data-ttu-id="b4dff-153">
  [Office.context.document.settings.set(name, value);](/javascript/api/office/office.settings#set-name--value-)</span><span class="sxs-lookup"><span data-stu-id="b4dff-153">[Office.context.document.settings.set(name, value);](/javascript/api/office/office.settings#set-name--value-)</span></span> | <span data-ttu-id="b4dff-154">設定はコンテンツ アドインによってのみサポートされます</span><span class="sxs-lookup"><span data-stu-id="b4dff-154">Settings are supported by content add-ins only</span></span> | 
| [<span data-ttu-id="b4dff-155">Office.EventType.DocumentSelectionChanged</span><span class="sxs-lookup"><span data-stu-id="b4dff-155">Office.EventType.DocumentSelectionChanged</span></span>](/javascript/api/office/office.documentselectionchangedeventargs) ||

<span data-ttu-id="b4dff-156">一般に、アプリケーション固有の API でサポートされていない操作を行う場合は、共通 API を使用します。</span><span class="sxs-lookup"><span data-stu-id="b4dff-156">In general, you use the Common API to do something that isn't supported in the application-specific API.</span></span> <span data-ttu-id="b4dff-157">共通 API の使用の詳細については、「[共通 JavaScript API オブジェクト モデル](../develop/office-javascript-api-object-model.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="b4dff-157">To learn more about using the Common API, see [Common JavaScript API object model](../develop/office-javascript-api-object-model.md).</span></span>

<a name="om-diagram"></a>
## <a name="onenote-object-model-diagram"></a><span data-ttu-id="b4dff-158">OneNote のオブジェクト モデル図</span><span class="sxs-lookup"><span data-stu-id="b4dff-158">OneNote object model diagram</span></span> 
<span data-ttu-id="b4dff-159">次の図では、OneNote JavaScript API で現在使用可能なものが示されます。</span><span class="sxs-lookup"><span data-stu-id="b4dff-159">The following diagram represents what's currently available in the OneNote JavaScript API.</span></span>

  ![OneNote のオブジェクト モデル図](../images/onenote-om.png)

## <a name="see-also"></a><span data-ttu-id="b4dff-161">関連項目</span><span class="sxs-lookup"><span data-stu-id="b4dff-161">See also</span></span>

- [<span data-ttu-id="b4dff-162">Office アドインを構築する</span><span class="sxs-lookup"><span data-stu-id="b4dff-162">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)
- [<span data-ttu-id="b4dff-163">最初の OneNote アドインをビルドする</span><span class="sxs-lookup"><span data-stu-id="b4dff-163">Build your first OneNote add-in</span></span>](../quickstarts/onenote-quickstart.md)
- [<span data-ttu-id="b4dff-164">OneNote JavaScript API リファレンス</span><span class="sxs-lookup"><span data-stu-id="b4dff-164">OneNote JavaScript API reference</span></span>](../reference/overview/onenote-add-ins-javascript-reference.md)
- [<span data-ttu-id="b4dff-165">Rubric Grader のサンプル</span><span class="sxs-lookup"><span data-stu-id="b4dff-165">Rubric Grader sample</span></span>](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [<span data-ttu-id="b4dff-166">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="b4dff-166">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
