---
title: OneNote の JavaScript API のプログラミングの概要
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: aded1210abc11a80c6200a207d3896df8ef4218b
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/23/2018
ms.locfileid: "19439013"
---
# <a name="onenote-javascript-api-programming-overview"></a><span data-ttu-id="348df-102">OneNote の JavaScript API のプログラミングの概要</span><span class="sxs-lookup"><span data-stu-id="348df-102">OneNote JavaScript API programming overview</span></span>

<span data-ttu-id="348df-103">OneNote では、OneNote Online アドインの JavaScript API が導入されています。OneNote オブジェクトを操作する作業ウィンドウ アドイン、コンテンツ アドイン、アドイン コマンドを作成し、Web サービスや他の Web ベースのリソースに接続できます。</span><span class="sxs-lookup"><span data-stu-id="348df-103">OneNote introduces a JavaScript API for OneNote Online add-ins. You can create task pane add-ins, content add-ins, and add-in commands that interact with OneNote objects and connect to web services or other web-based resources.</span></span>

> [!NOTE]
> <span data-ttu-id="348df-p101">AppSource にアドインを[公開](../publish/publish.md)し、Office エクスペリエンスで利用できるようにする予定がある場合は、[AppSource の検証ポリシー](https://docs.microsoft.com/en-us/office/dev/store/validation-policies)に準拠していることを確認してください。たとえば、検証に合格するには、定義したメソッドをサポートするすべてのプラットフォームでアドインが動作する必要があります (詳細については、[セクション 4.12](https://docs.microsoft.com/en-us/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) と [Office アドインを使用できるホストおよびプラットフォーム](../overview/office-add-in-availability.md)のページを参照してください)。</span><span class="sxs-lookup"><span data-stu-id="348df-p101">If you plan to [publish](../publish/publish.md) your add-in to AppSource and make it available within the Office experience, make sure that you conform to the [AppSource validation policies](https://docs.microsoft.com/en-us/office/dev/store/validation-policies). For example, to pass validation, your add-in must work across all platforms that support the methods that you define (for more information, see [section 4.12](https://docs.microsoft.com/en-us/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) and the [Office Add-in host and availability page](../overview/office-add-in-availability.md)).</span></span>

## <a name="components-of-an-office-add-in"></a><span data-ttu-id="348df-106">Office アドインのコンポーネント</span><span class="sxs-lookup"><span data-stu-id="348df-106">Components of an Office Add-in</span></span>

<span data-ttu-id="348df-107">アドインは、2 つの基本コンポーネントで構成されます。</span><span class="sxs-lookup"><span data-stu-id="348df-107">Add-ins consist of two basic components:</span></span>

- <span data-ttu-id="348df-p102">Web ページと必要な任意の JavaScript、CSS、他のファイルを含む **Web アプリケーション**。これらのファイルは、Web サーバーか、Microsoft Azure などの Web ホスティング サービスでホストされます。OneNote Online では、Web アプリケーションはブラウザー コントロールや iFrame で表示されます。</span><span class="sxs-lookup"><span data-stu-id="348df-p102">A **web application** consisting of a webpage and any required JavaScript, CSS, or other files. These files are hosted on a web server or web hosting service, such as Microsoft Azure. In OneNote Online, the web application displays in a browser control or iframe.</span></span>
    
- <span data-ttu-id="348df-p103">アドインの Web ページの URL とアドインの任意のアクセス要件、設定、機能を指定する **XML マニフェスト**。このファイルは、クライアントに保存されます。OneNote アドインは、他の Office アドインと同じ[マニフェスト](../develop/add-in-manifests.md)形式を使います。</span><span class="sxs-lookup"><span data-stu-id="348df-p103">An **XML manifest** that specifies the URL of the add-in's webpage and any access requirements, settings, and capabilities for the add-in. This file is stored on the client. OneNote add-ins use the same [manifest](../develop/add-in-manifests.md) format as other Office Add-ins.</span></span>

<span data-ttu-id="348df-114">**Office アドイン = マニフェスト + Web ページ**</span><span class="sxs-lookup"><span data-stu-id="348df-114">**Office Add-in = Manifest + Webpage**</span></span>

![Office アドインはマニフェストと Web ページによって構成されます](../images/onenote-add-in.png)

## <a name="using-the-javascript-api"></a><span data-ttu-id="348df-116">JavaScript API の使用</span><span class="sxs-lookup"><span data-stu-id="348df-116">Using the JavaScript API</span></span>

<span data-ttu-id="348df-p104">アドインは、ホスト アプリケーションのランタイム コンテキストを使って、JavaScript API にアクセスします。API には次の 2 つの階層があります。</span><span class="sxs-lookup"><span data-stu-id="348df-p104">Add-ins use the runtime context of the host application to access the JavaScript API. The API has two layers:</span></span> 

- <span data-ttu-id="348df-119">**アプリケーション** オブジェクトを通してアクセスされる、OneNote 固有の操作のための**豊富な API**。</span><span class="sxs-lookup"><span data-stu-id="348df-119">A **rich API** for OneNote-specific operations, accessed through the **Application** object.</span></span>
- <span data-ttu-id="348df-120">**ドキュメント** オブジェクトを通してアクセスされ、Office アプリケーション全体で共有される**共通 API**。</span><span class="sxs-lookup"><span data-stu-id="348df-120">A **common API** that's shared across Office applications, accessed through the **Document** object.</span></span>

### <a name="accessing-the-rich-api-through-the-application-object"></a><span data-ttu-id="348df-121">*アプリケーション* オブジェクトを使った豊富な API へのアクセス</span><span class="sxs-lookup"><span data-stu-id="348df-121">Accessing the rich API through the *Application* object</span></span>

<span data-ttu-id="348df-p105">**アプリケーション** オブジェクトを使って、**ノートブック**、**セクション**、**ページ**などの OneNote オブジェクトにアクセスします。豊富な API を使うと、プロキシ オブジェクトでバッチ操作を実行できます。基本的な流れは、以下のようになります。</span><span class="sxs-lookup"><span data-stu-id="348df-p105">Use the **Application** object to access OneNote objects such as **Notebook**, **Section**, and **Page**. With rich APIs, you run batch operations on proxy objects. The basic flow goes something like this:</span></span> 

1. <span data-ttu-id="348df-125">コンテキストからアプリケーション インスタンスを取得します。</span><span class="sxs-lookup"><span data-stu-id="348df-125">Get the application instance from the context.</span></span>

2. <span data-ttu-id="348df-p106">操作する OneNote オブジェクトを表すプロキシを作成します。プロキシ オブジェクトのプロパティの読み取りや書き込みを行い、メソッドを呼び出すことにより、プロキシ オブジェクトを同期的に操作します。</span><span class="sxs-lookup"><span data-stu-id="348df-p106">Create a proxy that represents the OneNote object you want to work with. You interact synchronously with proxy objects by reading and writing their properties and calling their methods.</span></span> 

3. <span data-ttu-id="348df-p107">プロキシで**読み込み**を呼び出し、パラメーターで指定されたプロパティ値を設定します。この呼び出しは、コマンドのキューに追加されます。</span><span class="sxs-lookup"><span data-stu-id="348df-p107">Call **load** on the proxy to fill it with the property values specified in the parameter. This call is added to the queue of commands.</span></span>

   > [!NOTE]
   > <span data-ttu-id="348df-130">API へのメソッドの呼び出し (`context.application.getActiveSection().pages;` など) も、キューに追加されます。</span><span class="sxs-lookup"><span data-stu-id="348df-130">Method calls to the API (such as `context.application.getActiveSection().pages;`) are also added to the queue.</span></span>

4. <span data-ttu-id="348df-p108">キューに置かれたすべてのコマンドをキューに置かれた順序で実行するには、**context.sync** を呼び出します。これにより、実行中のスクリプトと実際のオブジェクトの間の状態が同期されます。また、読み込まれた OneNote オブジェクトのプロパティを取得して、スクリプトで使います。追加のアクションのチェーン処理には、返された約束オブジェクトを使うことができます。</span><span class="sxs-lookup"><span data-stu-id="348df-p108">Call **context.sync** to run all queued commands in the order that they were queued. This synchronizes the state between your running script and the real objects, and by retrieving properties of loaded OneNote objects for use in your script. You can use the returned promise object for chaining additional actions.</span></span>

<span data-ttu-id="348df-134">例:</span><span class="sxs-lookup"><span data-stu-id="348df-134">For example:</span></span> 

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

<span data-ttu-id="348df-135">[API リファレンス](https://dev.office.com/reference/add-ins/onenote/onenote-add-ins-javascript-reference)では、サポートされている OneNote オブジェクトと操作を見つけることができます。</span><span class="sxs-lookup"><span data-stu-id="348df-135">You can find supported OneNote objects and operations in the [API reference](https://dev.office.com/reference/add-ins/onenote/onenote-add-ins-javascript-reference).</span></span>

### <a name="accessing-the-common-api-through-the-document-object"></a><span data-ttu-id="348df-136">*ドキュメント* オブジェクトを使った共通 API へのアクセス</span><span class="sxs-lookup"><span data-stu-id="348df-136">Accessing the common API through the *Document* object</span></span>

<span data-ttu-id="348df-137">**ドキュメント** オブジェクトを使って、[getSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.getselecteddataasync) メソッドや [setSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.setselecteddataasync) メソッドなどの共通 API にアクセスします。</span><span class="sxs-lookup"><span data-stu-id="348df-137">Use the **Document** object to access the common API, such as the [getSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.getselecteddataasync) and [setSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.setselecteddataasync) methods.</span></span> 

<span data-ttu-id="348df-138">例:</span><span class="sxs-lookup"><span data-stu-id="348df-138">For example:</span></span>  

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
<span data-ttu-id="348df-139">OneNote アドインは、次の共通 API のみをサポートします。</span><span class="sxs-lookup"><span data-stu-id="348df-139">OneNote add-ins support only the following common APIs:</span></span>

| <span data-ttu-id="348df-140">API</span><span class="sxs-lookup"><span data-stu-id="348df-140">API</span></span> | <span data-ttu-id="348df-141">メモ</span><span class="sxs-lookup"><span data-stu-id="348df-141">Notes</span></span> |
|:------|:------|
| [<span data-ttu-id="348df-142">Office.context.document.getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="348df-142">Office.context.document.getSelectedDataAsync</span></span>](https://dev.office.com/reference/add-ins/shared/document.getselecteddataasync) | <span data-ttu-id="348df-143">**Office.CoercionType.Text** と **Office.CoercionType.Matrix** のみ</span><span class="sxs-lookup"><span data-stu-id="348df-143">**Office.CoercionType.Text** and **Office.CoercionType.Matrix** only</span></span> |
| [<span data-ttu-id="348df-144">Office.context.document.setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="348df-144">Office.context.document.setSelectedDataAsync</span></span>](https://dev.office.com/reference/add-ins/shared/document.setselecteddataasync) | <span data-ttu-id="348df-145">**Office.CoercionType.Text**、**Office.CoercionType.Image**、**Office.CoercionType.Html** のみ</span><span class="sxs-lookup"><span data-stu-id="348df-145">**Office.CoercionType.Text**, **Office.CoercionType.Image**, and **Office.CoercionType.Html** only</span></span> | 
| [<span data-ttu-id="348df-146">var mySetting = Office.context.document.settings.get(name);</span><span class="sxs-lookup"><span data-stu-id="348df-146">var mySetting = Office.context.document.settings.get(name);</span></span>](https://dev.office.com/reference/add-ins/shared/settings.get) | <span data-ttu-id="348df-147">設定はコンテンツ アドインによってのみサポートされます</span><span class="sxs-lookup"><span data-stu-id="348df-147">Settings are supported by content add-ins only</span></span> | 
| [<span data-ttu-id="348df-148">Office.context.document.settings.set(name, value);</span><span class="sxs-lookup"><span data-stu-id="348df-148">Office.context.document.settings.set(name, value);</span></span>](https://dev.office.com/reference/add-ins/shared/settings.set) | <span data-ttu-id="348df-149">設定はコンテンツ アドインによってのみサポートされます</span><span class="sxs-lookup"><span data-stu-id="348df-149">Settings are supported by content add-ins only</span></span> | 
| [<span data-ttu-id="348df-150">Office.EventType.DocumentSelectionChanged</span><span class="sxs-lookup"><span data-stu-id="348df-150">Office.EventType.DocumentSelectionChanged</span></span>](https://dev.office.com/reference/add-ins/shared/document.selectionchanged.event) ||

<span data-ttu-id="348df-p109">一般に、豊富な API でサポートされていない操作を行う場合は、共通 API のみを使います。共通 API の使用について詳しくは、Office アドインの[ドキュメント](../overview/office-add-ins.md)と[リファレンス](https://dev.office.com/reference/add-ins/javascript-api-for-office)をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="348df-p109">In general, you only use the common API to do something that isn't supported in the rich API. To learn more about using the common API, see the Office Add-ins [documentation](../overview/office-add-ins.md) and [reference](https://dev.office.com/reference/add-ins/javascript-api-for-office).</span></span>


<a name="om-diagram"></a>
## <a name="onenote-object-model-diagram"></a><span data-ttu-id="348df-153">OneNote のオブジェクト モデル図</span><span class="sxs-lookup"><span data-stu-id="348df-153">OneNote object model diagram</span></span> 
<span data-ttu-id="348df-154">次の図では、OneNote JavaScript API で現在使用可能なものが示されます。</span><span class="sxs-lookup"><span data-stu-id="348df-154">The following diagram represents what's currently available in the OneNote JavaScript API.</span></span>

  ![OneNote のオブジェクト モデル図](../images/onenote-om.png)


## <a name="see-also"></a><span data-ttu-id="348df-156">関連項目</span><span class="sxs-lookup"><span data-stu-id="348df-156">See also</span></span>

- [<span data-ttu-id="348df-157">最初の OneNote アドインをビルドする</span><span class="sxs-lookup"><span data-stu-id="348df-157">Build your first OneNote add-in</span></span>](onenote-add-ins-getting-started.md)
- [<span data-ttu-id="348df-158">OneNote JavaScript API リファレンス</span><span class="sxs-lookup"><span data-stu-id="348df-158">OneNote JavaScript API reference</span></span>](https://dev.office.com/reference/add-ins/onenote/onenote-add-ins-javascript-reference)
- [<span data-ttu-id="348df-159">Rubric Grader のサンプル</span><span class="sxs-lookup"><span data-stu-id="348df-159">Rubric Grader sample</span></span>](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [<span data-ttu-id="348df-160">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="348df-160">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
