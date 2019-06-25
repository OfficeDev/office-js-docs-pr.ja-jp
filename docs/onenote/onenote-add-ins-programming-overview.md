---
title: OneNote の JavaScript API のプログラミングの概要
description: ''
ms.date: 06/20/2019
localization_priority: Priority
ms.openlocfilehash: 852c68bc9edf370d0eef687fb4869b23d4f59fe4
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/21/2019
ms.locfileid: "35128637"
---
# <a name="onenote-javascript-api-programming-overview"></a><span data-ttu-id="2cb26-102">OneNote の JavaScript API のプログラミングの概要</span><span class="sxs-lookup"><span data-stu-id="2cb26-102">OneNote JavaScript API programming overview</span></span>

<span data-ttu-id="2cb26-103">OneNote では、OneNote on the web アドインの JavaScript API が導入されています。</span><span class="sxs-lookup"><span data-stu-id="2cb26-103">OneNote introduces a JavaScript API for OneNote add-ins on the web.</span></span> <span data-ttu-id="2cb26-104">OneNote オブジェクトを操作する作業ウィンドウ アドイン、コンテンツ アドイン、アドイン コマンドを作成し、Web サービスやその他の Web ベースのリソースに接続できます。</span><span class="sxs-lookup"><span data-stu-id="2cb26-104">OneNote introduces a JavaScript API for OneNote Online add-ins. You can create task pane add-ins, content add-ins, and add-in commands that interact with OneNote objects and connect to web services or other web-based resources.</span></span>

> [!NOTE]
> <span data-ttu-id="2cb26-p102">AppSource にアドインを[公開](../publish/publish.md)し、Office エクスペリエンスで利用できるようにする予定がある場合は、[AppSource の検証ポリシー](/office/dev/store/validation-policies)に準拠していることを確認してください。たとえば、検証に合格するには、定義したメソッドをサポートするすべてのプラットフォームでアドインが動作する必要があります (詳細については、[セクション 4.12](/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) と [Office アドインを使用できるホストおよびプラットフォーム](../overview/office-add-in-availability.md)のページを参照してください)。</span><span class="sxs-lookup"><span data-stu-id="2cb26-p102">If you plan to [publish](../publish/publish.md) your add-in to AppSource and make it available within the Office experience, make sure that you conform to the [AppSource validation policies](/office/dev/store/validation-policies). For example, to pass validation, your add-in must work across all platforms that support the methods that you define (for more information, see [section 4.12](/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) and the [Office Add-in host and availability page](../overview/office-add-in-availability.md)).</span></span>

## <a name="components-of-an-office-add-in"></a><span data-ttu-id="2cb26-107">Office アドインのコンポーネント</span><span class="sxs-lookup"><span data-stu-id="2cb26-107">Components of an Office Add-in</span></span>

<span data-ttu-id="2cb26-108">アドインは、2 つの基本コンポーネントで構成されます。</span><span class="sxs-lookup"><span data-stu-id="2cb26-108">Add-ins consist of two basic components:</span></span>

- <span data-ttu-id="2cb26-109">Web ページと必要な任意の JavaScript、CSS、他のファイルで構成される **Web アプリケーション**。</span><span class="sxs-lookup"><span data-stu-id="2cb26-109">A **web application** consisting of a webpage and any required JavaScript, CSS, or other files.</span></span> <span data-ttu-id="2cb26-110">これらのファイルは、Web サーバーか、Microsoft Azure などの Web ホスティング サービスでホストされます。</span><span class="sxs-lookup"><span data-stu-id="2cb26-110">These files are hosted on a web server or web hosting service, such as Microsoft Azure.</span></span> <span data-ttu-id="2cb26-111">OneNote on the web では、Web アプリケーションはブラウザー コントロールや iFrame で表示されます。</span><span class="sxs-lookup"><span data-stu-id="2cb26-111">In OneNote Online, the web application displays in a browser control or iframe.</span></span>

- <span data-ttu-id="2cb26-p104">アドインの Web ページの URL とアドインの任意のアクセス要件、設定、機能を指定する **XML マニフェスト**。このファイルは、クライアントに保存されます。OneNote アドインは、他の Office アドインと同じ[マニフェスト](../develop/add-in-manifests.md)形式を使います。</span><span class="sxs-lookup"><span data-stu-id="2cb26-p104">An **XML manifest** that specifies the URL of the add-in's webpage and any access requirements, settings, and capabilities for the add-in. This file is stored on the client. OneNote add-ins use the same [manifest](../develop/add-in-manifests.md) format as other Office Add-ins.</span></span>

<span data-ttu-id="2cb26-115">**Office アドイン = マニフェスト + Web ページ**</span><span class="sxs-lookup"><span data-stu-id="2cb26-115">**Office Add-in = Manifest + Webpage**</span></span>

![Office アドインはマニフェストと Web ページによって構成されます](../images/onenote-add-in.png)

## <a name="using-the-javascript-api"></a><span data-ttu-id="2cb26-117">JavaScript API の使用</span><span class="sxs-lookup"><span data-stu-id="2cb26-117">Using the JavaScript API</span></span>

<span data-ttu-id="2cb26-p105">アドインは、ホスト アプリケーションのランタイム コンテキストを使って、JavaScript API にアクセスします。API には次の 2 つの階層があります。</span><span class="sxs-lookup"><span data-stu-id="2cb26-p105">Add-ins use the runtime context of the host application to access the JavaScript API. The API has two layers:</span></span> 

- <span data-ttu-id="2cb26-120">**アプリケーション** オブジェクトを通じてアクセスされる、OneNote 固有の操作のための**ホスト固有の API**。</span><span class="sxs-lookup"><span data-stu-id="2cb26-120">A **host-specific API** for OneNote-specific operations, accessed through the **Application** object.</span></span>
- <span data-ttu-id="2cb26-121">**ドキュメント** オブジェクトを通じてアクセスされ、Office アプリケーション全体で共有される**共通 API**。</span><span class="sxs-lookup"><span data-stu-id="2cb26-121">A **Common API** that's shared across Office applications, accessed through the **Document** object.</span></span>

### <a name="accessing-the-host-specific-api-through-the-application-object"></a><span data-ttu-id="2cb26-122">*アプリケーション* オブジェクトを使ったホスト固有の API へのアクセス</span><span class="sxs-lookup"><span data-stu-id="2cb26-122">Accessing the host-specific API through the *Application* object</span></span>

<span data-ttu-id="2cb26-123">**アプリケーション** オブジェクトを使って、**ノートブック**、**セクション**、**ページ**などの OneNote オブジェクトにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="2cb26-123">Use the **Application** object to access OneNote objects such as **Notebook**, **Section**, and **Page**.</span></span> <span data-ttu-id="2cb26-124">ホスト固有の API を使うと、プロキシ オブジェクトでバッチ操作を実行できます。</span><span class="sxs-lookup"><span data-stu-id="2cb26-124">With host-specific APIs, you run batch operations on proxy objects.</span></span> <span data-ttu-id="2cb26-125">基本的な流れは、以下のようになります。</span><span class="sxs-lookup"><span data-stu-id="2cb26-125">The basic flow goes something like this:</span></span> 

1. <span data-ttu-id="2cb26-126">コンテキストからアプリケーション インスタンスを取得します。</span><span class="sxs-lookup"><span data-stu-id="2cb26-126">Get the application instance from the context.</span></span>

2. <span data-ttu-id="2cb26-p107">操作する OneNote オブジェクトを表すプロキシを作成します。プロキシ オブジェクトのプロパティの読み取りや書き込みを行い、メソッドを呼び出すことにより、プロキシ オブジェクトを同期的に操作します。</span><span class="sxs-lookup"><span data-stu-id="2cb26-p107">Create a proxy that represents the OneNote object you want to work with. You interact synchronously with proxy objects by reading and writing their properties and calling their methods.</span></span>

3. <span data-ttu-id="2cb26-p108">プロキシで**読み込み**を呼び出し、パラメーターで指定されたプロパティ値を設定します。この呼び出しは、コマンドのキューに追加されます。</span><span class="sxs-lookup"><span data-stu-id="2cb26-p108">Call **load** on the proxy to fill it with the property values specified in the parameter. This call is added to the queue of commands.</span></span>

   > [!NOTE]
   > <span data-ttu-id="2cb26-131">API へのメソッドの呼び出し (`context.application.getActiveSection().pages;` など) も、キューに追加されます。</span><span class="sxs-lookup"><span data-stu-id="2cb26-131">Method calls to the API (such as `context.application.getActiveSection().pages;`) are also added to the queue.</span></span>

4. <span data-ttu-id="2cb26-p109">キューに置かれたすべてのコマンドをキューに置かれた順序で実行するには、**context.sync** を呼び出します。これにより、実行中のスクリプトと実際のオブジェクトの間の状態が同期されます。また、読み込まれた OneNote オブジェクトのプロパティを取得して、スクリプトで使います。追加のアクションのチェーン処理には、返された約束オブジェクトを使うことができます。</span><span class="sxs-lookup"><span data-stu-id="2cb26-p109">Call **context.sync** to run all queued commands in the order that they were queued. This synchronizes the state between your running script and the real objects, and by retrieving properties of loaded OneNote objects for use in your script. You can use the returned promise object for chaining additional actions.</span></span>

<span data-ttu-id="2cb26-135">例:</span><span class="sxs-lookup"><span data-stu-id="2cb26-135">For example:</span></span>

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

<span data-ttu-id="2cb26-136">[API リファレンス](/office/dev/add-ins/reference/overview/onenote-add-ins-javascript-reference)では、サポートされている OneNote オブジェクトと操作を見つけることができます。</span><span class="sxs-lookup"><span data-stu-id="2cb26-136">You can find supported OneNote objects and operations in the [API reference](/office/dev/add-ins/reference/overview/onenote-add-ins-javascript-reference).</span></span>

### <a name="accessing-the-common-api-through-the-document-object"></a><span data-ttu-id="2cb26-137">*ドキュメント* オブジェクトを使った共通 API へのアクセス</span><span class="sxs-lookup"><span data-stu-id="2cb26-137">Accessing the Common API through the *Document* object</span></span>

<span data-ttu-id="2cb26-138">**ドキュメント** オブジェクトを使って、[getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) メソッドや [setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) メソッドなどの共通 API にアクセスします。</span><span class="sxs-lookup"><span data-stu-id="2cb26-138">Use the **Document** object to access the Common API, such as the [getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) and [setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) methods.</span></span> 


<span data-ttu-id="2cb26-139">次に例を示します。</span><span class="sxs-lookup"><span data-stu-id="2cb26-139">For example:</span></span>  

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

<span data-ttu-id="2cb26-140">OneNote アドインは、次の共通 API のみをサポートします。</span><span class="sxs-lookup"><span data-stu-id="2cb26-140">OneNote add-ins support only the following Common APIs:</span></span>

| <span data-ttu-id="2cb26-141">API</span><span class="sxs-lookup"><span data-stu-id="2cb26-141">API</span></span> | <span data-ttu-id="2cb26-142">メモ</span><span class="sxs-lookup"><span data-stu-id="2cb26-142">Notes</span></span> |
|:------|:------|
| [<span data-ttu-id="2cb26-143">Office.context.document.getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="2cb26-143">Office.context.document.getSelectedDataAsync</span></span>](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) | <span data-ttu-id="2cb26-144">**Office.CoercionType.Text** と **Office.CoercionType.Matrix** のみ</span><span class="sxs-lookup"><span data-stu-id="2cb26-144">**Office.CoercionType.Text** and **Office.CoercionType.Matrix** only</span></span> |
| [<span data-ttu-id="2cb26-145">Office.context.document.setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="2cb26-145">Office.context.document.setSelectedDataAsync</span></span>](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) | <span data-ttu-id="2cb26-146">**Office.CoercionType.Text**、**Office.CoercionType.Image**、**Office.CoercionType.Html** のみ</span><span class="sxs-lookup"><span data-stu-id="2cb26-146">**Office.CoercionType.Text**, **Office.CoercionType.Image**, and **Office.CoercionType.Html** only</span></span> | 
| <span data-ttu-id="2cb26-147">
  [var mySetting = Office.context.document.settings.get(name);](/javascript/api/office/office.settings#get-name-)</span><span class="sxs-lookup"><span data-stu-id="2cb26-147">[var mySetting = Office.context.document.settings.get(name);](/javascript/api/office/office.settings#get-name-)</span></span> | <span data-ttu-id="2cb26-148">設定はコンテンツ アドインによってのみサポートされます</span><span class="sxs-lookup"><span data-stu-id="2cb26-148">Settings are supported by content add-ins only</span></span> | 
| <span data-ttu-id="2cb26-149">
  [Office.context.document.settings.set(name, value);](/javascript/api/office/office.settings#set-name--value-)</span><span class="sxs-lookup"><span data-stu-id="2cb26-149">[Office.context.document.settings.set(name, value);](/javascript/api/office/office.settings#set-name--value-)</span></span> | <span data-ttu-id="2cb26-150">設定はコンテンツ アドインによってのみサポートされます</span><span class="sxs-lookup"><span data-stu-id="2cb26-150">Settings are supported by content add-ins only</span></span> | 
| [<span data-ttu-id="2cb26-151">Office.EventType.DocumentSelectionChanged</span><span class="sxs-lookup"><span data-stu-id="2cb26-151">Office.EventType.DocumentSelectionChanged</span></span>](/javascript/api/office/office.documentselectionchangedeventargs) ||

<span data-ttu-id="2cb26-152">一般に、ホスト固有の API でサポートされていない操作を行う場合は、共通 API のみを使います。</span><span class="sxs-lookup"><span data-stu-id="2cb26-152">In general, you only use the Common API to do something that isn't supported in the host-specific API.</span></span> <span data-ttu-id="2cb26-153">共通 API の使用の詳細については、Office アドインの[ドキュメント](../overview/office-add-ins.md)と[リファレンス](../reference/javascript-api-for-office.md)をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="2cb26-153">To learn more about using the Common API, see the Office Add-ins [documentation](../overview/office-add-ins.md) and [reference](../reference/javascript-api-for-office.md).</span></span>


<a name="om-diagram"></a>
## <a name="onenote-object-model-diagram"></a><span data-ttu-id="2cb26-154">OneNote のオブジェクト モデル図</span><span class="sxs-lookup"><span data-stu-id="2cb26-154">OneNote object model diagram</span></span> 
<span data-ttu-id="2cb26-155">次の図では、OneNote JavaScript API で現在使用可能なものが示されます。</span><span class="sxs-lookup"><span data-stu-id="2cb26-155">The following diagram represents what's currently available in the OneNote JavaScript API.</span></span>

  ![OneNote のオブジェクト モデル図](../images/onenote-om.png)


## <a name="see-also"></a><span data-ttu-id="2cb26-157">関連項目</span><span class="sxs-lookup"><span data-stu-id="2cb26-157">See also</span></span>

- [<span data-ttu-id="2cb26-158">最初の OneNote アドインをビルドする</span><span class="sxs-lookup"><span data-stu-id="2cb26-158">Build your first OneNote add-in</span></span>](../quickstarts/onenote-quickstart.md)
- [<span data-ttu-id="2cb26-159">OneNote JavaScript API リファレンス</span><span class="sxs-lookup"><span data-stu-id="2cb26-159">OneNote JavaScript API reference</span></span>](/office/dev/add-ins/reference/overview/onenote-add-ins-javascript-reference)
- [<span data-ttu-id="2cb26-160">Rubric Grader のサンプル</span><span class="sxs-lookup"><span data-stu-id="2cb26-160">Rubric Grader sample</span></span>](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [<span data-ttu-id="2cb26-161">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="2cb26-161">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
