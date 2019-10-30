---
title: PowerPoint アドインのチュートリアル
description: このチュートリアルでは、画像の挿入、テキストの挿入、スライドのメタデータ取得、およびスライド間の移動のための PowerPoint アドインを作成します。
ms.date: 10/29/2019
ms.prod: powerpoint
localization_priority: Normal
ms.openlocfilehash: 73d7e041a10a3991d2ba87b420eece191603983a
ms.sourcegitcommit: 818036a7163b1513d047e66a20434060415df241
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/29/2019
ms.locfileid: "37775299"
---
# <a name="tutorial-create-a-powerpoint-task-pane-add-in"></a><span data-ttu-id="85125-103">チュートリアル: PowerPoint 作業ウィンドウ アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="85125-103">Tutorial: Create a PowerPoint task pane add-in</span></span>

<span data-ttu-id="85125-104">このチュートリアルでは Visual Studio を使用して、以下を実行する PowerPoint 作業ウィンドウ アドインを作成します。</span><span class="sxs-lookup"><span data-stu-id="85125-104">In this tutorial, you'll use Visual Studio to create an PowerPoint task pane add-in that:</span></span>

> [!div class="checklist"]
> * <span data-ttu-id="85125-105">その日の [Bing](https://www.bing.com) の写真をスライドに追加する</span><span class="sxs-lookup"><span data-stu-id="85125-105">Adds the [Bing](https://www.bing.com) photo of the day to a slide</span></span>
> * <span data-ttu-id="85125-106">スライドにテキストを追加する</span><span class="sxs-lookup"><span data-stu-id="85125-106">Adds text to a slide</span></span>
> * <span data-ttu-id="85125-107">スライドのメタデータを取得する</span><span class="sxs-lookup"><span data-stu-id="85125-107">Gets slide metadata</span></span>
> * <span data-ttu-id="85125-108">スライド間を移動する</span><span class="sxs-lookup"><span data-stu-id="85125-108">Navigates between slides</span></span>

## <a name="prerequisites"></a><span data-ttu-id="85125-109">前提条件</span><span class="sxs-lookup"><span data-stu-id="85125-109">Prerequisites</span></span>

[!include[Quick Start prerequisites](../includes/quickstart-vs-prerequisites.md)]

## <a name="create-your-add-in-project"></a><span data-ttu-id="85125-110">アドイン プロジェクトの作成</span><span class="sxs-lookup"><span data-stu-id="85125-110">Create your add-in project</span></span>

<span data-ttu-id="85125-111">Visual Studio を使用して PowerPoint アドイン プロジェクトを作成するには次の手順を実行します。</span><span class="sxs-lookup"><span data-stu-id="85125-111">Complete the following steps to create a PowerPoint add-in project using Visual Studio.</span></span>

1. <span data-ttu-id="85125-112">[**新規プロジェクトの作成**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="85125-112">Choose **Create a new project**.</span></span>

2. <span data-ttu-id="85125-113">検索ボックスを使用して、**アドイン**と入力します。</span><span class="sxs-lookup"><span data-stu-id="85125-113">Using the search box, enter **add-in**.</span></span> <span data-ttu-id="85125-114">[**PowerPoint Web アドイン**] を選択し、[**次へ**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="85125-114">Choose **PowerPoint Web Add-in**, then select **Next**.</span></span>

3. <span data-ttu-id="85125-115">プロジェクト`HelloWorld`に名前を指定し、[**作成**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="85125-115">Name the project `HelloWorld`, and select **Create**.</span></span>

4. <span data-ttu-id="85125-116">**[Office アドインの作成]** ダイアログ ウィンドウで、**[新機能を PowerPoint に追加する]** を選択してから、**[完了]** を選択してプロジェクトを作成します。</span><span class="sxs-lookup"><span data-stu-id="85125-116">In the **Create Office Add-in** dialog window, choose **Add new functionalities to PowerPoint**, and then choose **Finish** to create the project.</span></span>

5. <span data-ttu-id="85125-p102">Visual Studio によってソリューションとその 2 つのプロジェクトが作成され、**ソリューション エクスプローラー**に表示されます。**Home.html** ファイルが Visual Studio で開かれます。</span><span class="sxs-lookup"><span data-stu-id="85125-p102">Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.</span></span>

     ![PowerPoint チュートリアル - HelloWorld ソリューションで 2 つのプロジェクトを表示する Visual Studio ソリューション エクスプローラー ウィンドウ](../images/powerpoint-tutorial-solution-explorer.png)

### <a name="explore-the-visual-studio-solution"></a><span data-ttu-id="85125-120">Visual Studio ソリューションについて理解する</span><span class="sxs-lookup"><span data-stu-id="85125-120">Explore the Visual Studio solution</span></span>

[!include[Description of Visual Studio projects](../includes/quickstart-vs-solution.md)]

### <a name="update-code"></a><span data-ttu-id="85125-121">コードを更新する</span><span class="sxs-lookup"><span data-stu-id="85125-121">Update code</span></span> 

<span data-ttu-id="85125-122">アドイン コードを次のように編集し、このチュートリアルの後続の手順でアドイン機能を実装するために使用するフレームワークを作成します。</span><span class="sxs-lookup"><span data-stu-id="85125-122">Edit the add-in code as follows to create the framework that you'll use to implement add-in functionality in subsequent steps of this tutorial.</span></span>

1. <span data-ttu-id="85125-p103">**Home.html** では、アドインの作業ウィンドウにレンダリングされる HTML を指定します。 **Home.html** において、\*\*\*\* で `id="content-main"` を検索し、**div** 全体を次のマークアップと置き換えてファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="85125-p103">**Home.html** specifies the HTML that will be rendered in the add-in's task pane. In **Home.html**, find the **div** with `id="content-main"`, replace that entire **div** with the following markup, and save the file.</span></span>

    ```html
    <!-- TODO2: Create the content-header div. -->
    <div id="content-main">
        <div class="padding">
            <!-- TODO1: Create the insert-image button. -->
            <!-- TODO3: Create the insert-text button. -->
            <!-- TODO4: Create the get-slide-metadata button. -->
            <!-- TODO5: Create the go-to-slide buttons. -->
        </div>
    </div>
    ```

2. <span data-ttu-id="85125-p104">Web アプリケーション プロジェクトのルートにあるファイル **Home.js** を開きます。 このファイルは、アドイン用のスクリプトを指定します。 すべての内容を次のコードに置き換え、ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="85125-p104">Open the file **Home.js** in the root of the web application project. This file specifies the script for the add-in. Replace the entire contents with the following code and save the file.</span></span>

    ```js
    (function () {
        "use strict";

        var messageBanner;

        Office.onReady(function () {
            $(document).ready(function () {
                // Initialize the FabricUI notification mechanism and hide it
                var element = document.querySelector('.MessageBanner');
                messageBanner = new components.MessageBanner(element);
                messageBanner.hideBanner();

                // TODO1: Assign event handler for insert-image button.
                // TODO4: Assign event handler for insert-text button.
                // TODO6: Assign event handler for get-slide-metadata button.
                // TODO8: Assign event handlers for the four navigation buttons.
            });
        });

        // TODO2: Define the insertImage function. 

        // TODO3: Define the insertImageFromBase64String function.

        // TODO5: Define the insertText function.

        // TODO7: Define the getSlideMetadata function.

        // TODO9: Define the navigation functions.

        // Helper function for displaying notifications
        function showNotification(header, content) {
            $("#notification-header").text(header);
            $("#notification-body").text(content);
            messageBanner.showBanner();
            messageBanner.toggleExpansion();
        }
    })();
    ```

## <a name="insert-an-image"></a><span data-ttu-id="85125-128">画像の挿入</span><span class="sxs-lookup"><span data-stu-id="85125-128">Insert an image</span></span>

<span data-ttu-id="85125-129">その日の [Bing](https://www.bing.com) の写真取得し、その画像をスライドに挿入するコードを追加するには、次の手順を実行します。</span><span class="sxs-lookup"><span data-stu-id="85125-129">Complete the following steps to add code that retrieves the [Bing](https://www.bing.com) photo of the day and inserts that image into a slide.</span></span>

1. <span data-ttu-id="85125-130">ソリューション エクスプローラーを使用して、**Controllers** という名前の新しいフォルダーを **HelloWorldWeb** プロジェクトに追加します。</span><span class="sxs-lookup"><span data-stu-id="85125-130">Using Solution Explorer, add a new folder named **Controllers** to the **HelloWorldWeb** project.</span></span>

    ![PowerPoint のチュートリアル - HelloWorldWeb プロジェクトの Controllers フォルダーを強調表示する Visual Studio ソリューション エクスプローラー ウィンドウ](../images/powerpoint-tutorial-solution-explorer-controllers.png)

2. <span data-ttu-id="85125-132">**Controllers** フォルダーを右クリックし、**[追加] > [新規スキャフォールディング アイテム...]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="85125-132">Right-click the **Controllers** folder and select **Add > New Scaffolded Item...**.</span></span>

3. <span data-ttu-id="85125-133">**[スキャフォールディングを追加]** ダイアログ ウィンドウで、「**Web API 2 Controller - Empty**」を選択し、**[追加]** ボタンを選択します。</span><span class="sxs-lookup"><span data-stu-id="85125-133">In the **Add Scaffold** dialog window, select **Web API 2 Controller - Empty** and choose the **Add** button.</span></span> 

4. <span data-ttu-id="85125-p105">**[コントローラーの追加]** ダイアログ ウィンドウでコントローラー名として「**PhotoController**」と入力し、**[追加]** ボタンを選択します。 Visual Studio によって **PhotoController.cs** ファイルが作成され、表示されます。</span><span class="sxs-lookup"><span data-stu-id="85125-p105">In the **Add Controller** dialog window, enter **PhotoController** as the controller name and choose the **Add** button. Visual Studio creates and opens the **PhotoController.cs** file.</span></span>

5. <span data-ttu-id="85125-p106">**PhotoController.cs** ファイルの内容全体を、Bing サービスを呼び出す次のコードに置き換え、その日の写真を Base64 でエンコードされた文字列として取得します。 Office JavaScript API を使用してイメージをドキュメントに挿入する場合は、イメージ データを Base64 でエンコードされた文字列として指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="85125-p106">Replace the entire contents of the **PhotoController.cs** file with the following code that calls the Bing service to retrieve the photo of the day as a Base64 encoded string. When you use the Office JavaScript API to insert an image into a document, the image data must be specified as a Base64 encoded string.</span></span>

    ```csharp
    using System;
    using System.IO;
    using System.Net;
    using System.Text;
    using System.Web.Http;
    using System.Xml;

    namespace HelloWorldWeb.Controllers
    {
        public class PhotoController : ApiController
        {
            public string Get()
            {
                string url = "http://www.bing.com/HPImageArchive.aspx?format=xml&idx=0&n=1";

                // Create the request.
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
                WebResponse response = request.GetResponse();

                using (Stream responseStream = response.GetResponseStream())
                {
                    // Process the result.
                    StreamReader reader = new StreamReader(responseStream, Encoding.UTF8);
                    string result = reader.ReadToEnd();

                    // Parse the xml response and to get the URL.
                    XmlDocument doc = new XmlDocument();
                    doc.LoadXml(result);
                    string photoURL = "http://bing.com" + doc.SelectSingleNode("/images/image/url").InnerText;

                    // Fetch the photo and return it as a Base64 encoded string.
                    return getPhotoFromURL(photoURL);
                }
            }

            private string getPhotoFromURL(string imageURL)
            {
                var webClient = new WebClient();
                byte[] imageBytes = webClient.DownloadData(imageURL);
                return Convert.ToBase64String(imageBytes);
            }
        }
    }
    ```

6. <span data-ttu-id="85125-p107">**Home.html** ファイルで `TODO1` を次のマークアップに置き換えます。 このマークアップにより、アドインの作業ウィンドウ内に表示される **[イメージの挿入]** ボタンを定義します。</span><span class="sxs-lookup"><span data-stu-id="85125-p107">In the **Home.html** file, replace `TODO1` with the following markup. This markup defines the **Insert Image** button that will appear within the add-in's task pane.</span></span>

    ```html
    <button class="Button Button--primary" id="insert-image">
        <span class="Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="Button-label">Insert Image</span>
        <span class="Button-description">Gets the photo of the day that shows on the Bing home page and adds it to the slide.</span>
    </button>
    ```

7. <span data-ttu-id="85125-140">**Home.js** ファイルで `TODO1` を次のコードに置き換え、**[イメージの挿入]** ボタンのイベント ハンドラーを割り当てます。</span><span class="sxs-lookup"><span data-stu-id="85125-140">In the **Home.js** file, replace `TODO1` with the following code to assign the event handler for the **Insert Image** button.</span></span>

    ```js
    $('#insert-image').click(insertImage);
    ```

8. <span data-ttu-id="85125-p108">**Home.js** ファイルで `TODO2` を次のコードに置き換え、**insertImage** 関数を定義します。 この関数は Bing Web サービスからイメージをフェッチし、`insertImageFromBase64String` 関数を呼び出してそのイメージをドキュメントに挿入します。</span><span class="sxs-lookup"><span data-stu-id="85125-p108">In the **Home.js** file, replace `TODO2` with the following code to define the **insertImage** function. This function fetches the image from the Bing web service and then calls the `insertImageFromBase64String` function to insert that image into the document.</span></span>

    ```js
    function insertImage() {
        // Get image from from web service (as a Base64 encoded string).
        $.ajax({
            url: "/api/Photo/", success: function (result) {
                insertImageFromBase64String(result);
            }, error: function (xhr, status, error) {
                showNotification("Error", "Oops, something went wrong.");
            }
        });
    }
    ```

9. <span data-ttu-id="85125-p109">**Home.js** ファイルで `TODO3` を次のコードに置き換え、`insertImageFromBase64String` 関数を定義します。 この関数は Office JavaScript API を使用してイメージをドキュメントに挿入します。 注意:</span><span class="sxs-lookup"><span data-stu-id="85125-p109">In the **Home.js** file, replace `TODO3` with the following code to define the `insertImageFromBase64String` function. This function uses the Office JavaScript API to insert the image into the document. Note:</span></span> 

    - <span data-ttu-id="85125-146">`coercionType` 要求の 2 番目のパラメーターとして指定されている `setSelectedDataAsyc` オプションは、挿入されるデータの種類を示します。</span><span class="sxs-lookup"><span data-stu-id="85125-146">The `coercionType` option that's specified as the second parameter of the `setSelectedDataAsyc` request indicates the type of data being inserted.</span></span> 

    - <span data-ttu-id="85125-147">`asyncResult` オブジェクトは `setSelectedDataAsync` 要求が失敗した場合の状態やエラー情報など、その要求の結果をカプセル化します。</span><span class="sxs-lookup"><span data-stu-id="85125-147">The `asyncResult` object encapsulates the result of the `setSelectedDataAsync` request, including status and error information if the request failed.</span></span>

    ```js
    function insertImageFromBase64String(image) {
        // Call Office.js to insert the image into the document.
        Office.context.document.setSelectedDataAsync(image, {
            coercionType: Office.CoercionType.Image
        },
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }
    ```

### <a name="test-the-add-in"></a><span data-ttu-id="85125-148">アドインをテストする</span><span class="sxs-lookup"><span data-stu-id="85125-148">Test the add-in</span></span>

1. <span data-ttu-id="85125-p110">Visual Studio を使用して、新しく作成した PowerPoint アドインをテストします。そのために、**F5** キーを押すか **[開始]** ボタンをクリックして、リボンに **[作業ウィンドウの表示]** アドイン ボタンが表示された PowerPoint を起動します。アドインは IIS 上でローカルにホストされます。</span><span class="sxs-lookup"><span data-stu-id="85125-p110">Using Visual Studio, test the newly created PowerPoint add-in by pressing **F5** or choosing the **Start** button to launch PowerPoint with the **Show Taskpane** add-in button displayed in the ribbon. The add-in will be hosted locally on IIS.</span></span>

    ![[開始] ボタンが強調表示されている Visual Studio のスクリーンショット](../images/powerpoint-tutorial-start.png)

2. <span data-ttu-id="85125-152">PowerPoint でリボンの **[作業ウィンドウの表示]** ボタンをクリックし、アドインの作業ウィンドウを開きます。</span><span class="sxs-lookup"><span data-stu-id="85125-152">In PowerPoint, select the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![[ホーム] リボンで [作業ウィンドウの表示] ボタンが強調表示されている Visual Studio のスクリーンショット](../images/powerpoint-tutorial-show-taskpane-button.png)

3. <span data-ttu-id="85125-154">作業ウィンドウで、**[イメージの挿入]** ボタンを押してその日の Bing 写真を現在のスライドに追加します。</span><span class="sxs-lookup"><span data-stu-id="85125-154">In the task pane, choose the **Insert Image** button to add the Bing photo of the day to the current slide.</span></span>

    ![[イメージの挿入] ボタンが強調表示されている PowerPoint アドインのスクリーンショット](../images/powerpoint-tutorial-insert-image-button.png)

4. <span data-ttu-id="85125-156">Visual Studio で **Shift + F5** を押すか **[停止]** ボタンを選択してアドインを停止します。</span><span class="sxs-lookup"><span data-stu-id="85125-156">In Visual Studio, stop the add-in by pressing **Shift + F5** or choosing the **Stop** button.</span></span> <span data-ttu-id="85125-157">アドインが停止すると、PowerPoint は自動的に閉じます。</span><span class="sxs-lookup"><span data-stu-id="85125-157">PowerPoint will automatically close when the add-in is stopped.</span></span>

    ![[停止] ボタンが強調表示されている Visual Studio のスクリーンショット](../images/powerpoint-tutorial-stop.png)

## <a name="customize-user-interface-ui-elements"></a><span data-ttu-id="85125-159">ユーザー インターフェイス (UI) 要素のカスタマイズ</span><span class="sxs-lookup"><span data-stu-id="85125-159">Customize User Interface (UI) elements</span></span>

<span data-ttu-id="85125-160">作業ウィンドウの UI をカスタマイズするマークアップを追加するには、次の手順を実行します。</span><span class="sxs-lookup"><span data-stu-id="85125-160">Complete the following steps to add markup that customizes the task pane UI.</span></span>

1. <span data-ttu-id="85125-p112">**Home.html** ファイルで `TODO2` を次のマークアップと置き換え、ヘッダー セクションとタイトルを作業ウィンドウに追加します。 注意:</span><span class="sxs-lookup"><span data-stu-id="85125-p112">In the **Home.html** file, replace `TODO2` with the following markup to add a header section and title to the task pane. Note:</span></span>

    - <span data-ttu-id="85125-p113">`ms-` で始まるスタイルは、[Office UI Fabric](../design/office-ui-fabric.md) で定義されています。これは、Office と Office 365 のユーザー エクスペリエンスを構築するための JavaScript フロント エンドのフレームワークです。 **Home.html** ファイルには、Fabric スタイル シートへの参照が含まれています。</span><span class="sxs-lookup"><span data-stu-id="85125-p113">The styles that begin with `ms-` are defined by [Office UI Fabric](../design/office-ui-fabric.md), a JavaScript front-end framework for building user experiences for Office and Office 365. The **Home.html** file includes a reference to the Fabric stylesheet.</span></span>

    ```html
    <div id="content-header">
        <div class="ms-Grid ms-bgColor-neutralPrimary">
            <div class="ms-Grid-row">
                <div class="padding ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12"> <div class="ms-font-xl ms-fontColor-white ms-fontWeight-semibold">My PowerPoint add-in</div></div>
            </div>
        </div>
    </div>
    ```

2. <span data-ttu-id="85125-165">**Home.html** ファイルにおいて、`class="footer"` で **div** を検索し、**div** 全体を削除して作業ウィンドウからフッター セクションを削除します。</span><span class="sxs-lookup"><span data-stu-id="85125-165">In the **Home.html** file, find the **div** with `class="footer"` and delete that entire **div** to remove the footer section from the task pane.</span></span>

### <a name="test-the-add-in"></a><span data-ttu-id="85125-166">アドインをテストする</span><span class="sxs-lookup"><span data-stu-id="85125-166">Test the add-in</span></span>

1. <span data-ttu-id="85125-167">Visual Studio を使用して、PowerPoint アドインをテストします。そのために、**F5** キーを押すか **[開始]** ボタンをクリックして、リボンに **[作業ウィンドウの表示]** アドイン ボタンが表示された PowerPoint を起動します。</span><span class="sxs-lookup"><span data-stu-id="85125-167">Using Visual Studio, test the PowerPoint add-in by pressing **F5** or choosing the **Start** button to launch PowerPoint with the **Show Taskpane** add-in button displayed in the ribbon.</span></span> <span data-ttu-id="85125-168">アドインは IIS 上でローカルにホストされます。</span><span class="sxs-lookup"><span data-stu-id="85125-168">The add-in will be hosted locally on IIS.</span></span>

    ![[開始] ボタンが強調表示されている Visual Studio のスクリーンショット](../images/powerpoint-tutorial-start.png)

2. <span data-ttu-id="85125-170">PowerPoint でリボンの **[作業ウィンドウの表示]** ボタンをクリックし、アドインの作業ウィンドウを開きます。</span><span class="sxs-lookup"><span data-stu-id="85125-170">In PowerPoint, select the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![[ホーム] リボンで [作業ウィンドウの表示] ボタンが強調表示されている Visual Studio のスクリーンショット](../images/powerpoint-tutorial-show-taskpane-button.png)

3. <span data-ttu-id="85125-172">このとき、作業ウィンドウにはヘッダー セクションとタイトルが含まれ、フッター セクションが含まれないことがわかります。</span><span class="sxs-lookup"><span data-stu-id="85125-172">Notice that the task pane now contains a header section and title, and no longer contains a footer section.</span></span>

    ![[画像の挿入] ボタンが強調表示されている PowerPoint アドインのスクリーンショット](../images/powerpoint-tutorial-new-task-pane-ui.png)

4. <span data-ttu-id="85125-174">Visual Studio で **Shift + F5** を押すか **[停止]** ボタンを選択してアドインを停止します。</span><span class="sxs-lookup"><span data-stu-id="85125-174">In Visual Studio, stop the add-in by pressing **Shift + F5** or choosing the **Stop** button.</span></span> <span data-ttu-id="85125-175">アドインが停止すると、PowerPoint は自動的に閉じます。</span><span class="sxs-lookup"><span data-stu-id="85125-175">PowerPoint will automatically close when the add-in is stopped.</span></span>

    ![[停止] ボタンが強調表示されている Visual Studio のスクリーンショット](../images/powerpoint-tutorial-stop.png)

## <a name="insert-text"></a><span data-ttu-id="85125-177">テキストの挿入</span><span class="sxs-lookup"><span data-stu-id="85125-177">Insert text</span></span>

<span data-ttu-id="85125-178">その日の [Bing](https://www.bing.com) の写真を含むタイトル スライドにテキストを挿入するコードを追加するには、次の手順を実行します。</span><span class="sxs-lookup"><span data-stu-id="85125-178">Complete the following steps to add code that inserts text into the title slide which contains the [Bing](https://www.bing.com) photo of the day.</span></span>

1. <span data-ttu-id="85125-p116">**Home.html** ファイルで `TODO3` を次のマークアップに置き換えます。 このマークアップにより、アドインの作業ウィンドウ内に表示される **[テキストの挿入]** ボタンを定義します。</span><span class="sxs-lookup"><span data-stu-id="85125-p116">In the **Home.html** file, replace `TODO3` with the following markup. This markup defines the **Insert Text** button that will appear within the add-in's task pane.</span></span>

    ```html
        <br /><br />
        <button class="Button Button--primary" id="insert-text">
            <span class="Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
            <span class="Button-label">Insert Text</span>
            <span class="Button-description">Inserts text into the slide.</span>
        </button>
    ```

2. <span data-ttu-id="85125-181">**Home.js** ファイルで `TODO4` を次のコードに置き換え、**[テキストの挿入]** ボタンのイベント ハンドラーを割り当てます。</span><span class="sxs-lookup"><span data-stu-id="85125-181">In the **Home.js** file, replace `TODO4` with the following code to assign the event handler for the **Insert Text** button.</span></span>

    ```js
    $('#insert-text').click(insertText);
    ```

3. <span data-ttu-id="85125-p117">**Home.js** ファイルで `TODO5` を次のコードに置き換え、**insertText** 関数を定義します。 この関数は、現在のスライドにテキストを挿入します。</span><span class="sxs-lookup"><span data-stu-id="85125-p117">In the **Home.js** file, replace `TODO5` with the following code to define the **insertText** function. This function inserts text into the current slide.</span></span>

    ```js
    function insertText() {
        Office.context.document.setSelectedDataAsync('Hello World!',
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }
    ```

### <a name="test-the-add-in"></a><span data-ttu-id="85125-184">アドインをテストする</span><span class="sxs-lookup"><span data-stu-id="85125-184">Test the add-in</span></span>

1. <span data-ttu-id="85125-185">Visual Studio を使用して、アドインをテストします。そのために、**F5** キーを押すか **[開始]** ボタンをクリックして、リボンに **[作業ウィンドウの表示]** アドイン ボタンが表示された PowerPoint を起動します。</span><span class="sxs-lookup"><span data-stu-id="85125-185">Using Visual Studio, test the add-in by pressing **F5** or choosing the **Start** button to launch PowerPoint with the **Show Taskpane** add-in button displayed in the ribbon.</span></span> <span data-ttu-id="85125-186">アドインは IIS 上でローカルにホストされます。</span><span class="sxs-lookup"><span data-stu-id="85125-186">The add-in will be hosted locally on IIS.</span></span>

    ![[開始] ボタンが強調表示されている Visual Studio のスクリーンショット](../images/powerpoint-tutorial-start.png)

2. <span data-ttu-id="85125-188">PowerPoint でリボンの **[作業ウィンドウの表示]** ボタンをクリックし、アドインの作業ウィンドウを開きます。</span><span class="sxs-lookup"><span data-stu-id="85125-188">In PowerPoint, select the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![[ホーム] リボンで [作業ウィンドウの表示] ボタンが強調表示されている Visual Studio のスクリーンショット](../images/powerpoint-tutorial-show-taskpane-button.png)

3. <span data-ttu-id="85125-190">作業ウィンドウで **[イメージの挿入]** ボタンをクリックしてその日の Bing 写真を現在のスライドに追加し、そのタイトルにテキスト ボックスが含まれるデザインをそのスライドに選択します。</span><span class="sxs-lookup"><span data-stu-id="85125-190">In the task pane, choose the **Insert Image** button to add the Bing photo of the day to the current slide and choose a design for the slide that contains a text box for the title.</span></span>

    ![[イメージの挿入] ボタンが強調表示されている PowerPoint アドインのスクリーンショット](../images/powerpoint-tutorial-insert-image-slide-design.png)

4. <span data-ttu-id="85125-192">タイトル スライドのテキスト ボックスにカーソルを置き、作業ウィンドウで **[テキストの挿入]** ボタンをクリックしてテキストをスライドに追加します。</span><span class="sxs-lookup"><span data-stu-id="85125-192">Put your cursor in the text box on the title slide and then in the task pane, choose the **Insert Text** button to add text to the slide.</span></span>

    ![[テキストの挿入] ボタンが強調表示されている PowerPoint アドインのスクリーンショット](../images/powerpoint-tutorial-insert-text.png)


5. <span data-ttu-id="85125-194">Visual Studio で **Shift + F5** を押すか **[停止]** ボタンを選択してアドインを停止します。</span><span class="sxs-lookup"><span data-stu-id="85125-194">In Visual Studio, stop the add-in by pressing **Shift + F5** or choosing the **Stop** button.</span></span> <span data-ttu-id="85125-195">アドインが停止すると、PowerPoint は自動的に閉じます。</span><span class="sxs-lookup"><span data-stu-id="85125-195">PowerPoint will automatically close when the add-in is stopped.</span></span>

    ![[停止] ボタンが強調表示されている Visual Studio のスクリーンショット](../images/powerpoint-tutorial-stop.png)

## <a name="get-slide-metadata"></a><span data-ttu-id="85125-197">スライドのメタデータの取得</span><span class="sxs-lookup"><span data-stu-id="85125-197">Get slide metadata</span></span>

<span data-ttu-id="85125-198">選択したスライドのメタデータを取得するコードを追加するには、次の手順を実行します。</span><span class="sxs-lookup"><span data-stu-id="85125-198">Complete the following steps to add code that retrieves metadata for the selected slide.</span></span>

1. <span data-ttu-id="85125-p120">**Home.html** ファイルで `TODO4` を次のマークアップに置き換えます。 このマークアップにより、アドインの作業ウィンドウ内に表示される **[Get Slide Metadata]** (スライドのメタデータの取得) ボタンを定義します。</span><span class="sxs-lookup"><span data-stu-id="85125-p120">In the **Home.html** file, replace `TODO4` with the following markup. This markup defines the **Get Slide Metadata** button that will appear within the add-in's task pane.</span></span>

    ```html
    <br /><br />
    <button class="Button Button--primary" id="get-slide-metadata">
        <span class="Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="Button-label">Get Slide Metadata</span>
        <span class="Button-description">Gets metadata for the selected slide(s).</span>
    </button>
    ```

2. <span data-ttu-id="85125-201">**Home.js** ファイルで `TODO6` を次のコードに置き換え、**[Get Slide Metadata]** (スライドのメタデータの取得) ボタンのイベント ハンドラーを割り当てます。</span><span class="sxs-lookup"><span data-stu-id="85125-201">In the **Home.js** file, replace `TODO6` with the following code to assign the event handler for the **Get Slide Metadata** button.</span></span>

    ```js
    $('#get-slide-metadata').click(getSlideMetadata);
    ```

3. <span data-ttu-id="85125-p121">**Home.js** ファイルで `TODO7` を次のコードに置き換え、**getSlideMetadata** 関数を定義します。 この関数は選択したスライドのメタデータを取得し、それをアドインの作業ウィンドウ内のポップアップ ダイアログ ウィンドウに書き込みます。</span><span class="sxs-lookup"><span data-stu-id="85125-p121">In the **Home.js** file, replace `TODO7` with the following code to define the **getSlideMetadata** function. This function retrieves metadata for the selected slide(s) and writes it to a popup dialog window within the add-in task pane.</span></span>

    ```js
    function getSlideMetadata() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange,
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    showNotification("Error", asyncResult.error.message);
                } else {
                    showNotification("Metadata for selected slide(s):", JSON.stringify(asyncResult.value), null, 2);
                }
            }
        );
    }
    ```

### <a name="test-the-add-in"></a><span data-ttu-id="85125-204">アドインをテストする</span><span class="sxs-lookup"><span data-stu-id="85125-204">Test the add-in</span></span>

1. <span data-ttu-id="85125-205">Visual Studio を使用して、アドインをテストします。そのために、**F5** キーを押すか **[開始]** ボタンをクリックして、リボンに **[作業ウィンドウの表示]** アドイン ボタンが表示された PowerPoint を起動します。</span><span class="sxs-lookup"><span data-stu-id="85125-205">Using Visual Studio, test the add-in by pressing **F5** or choosing the **Start** button to launch PowerPoint with the **Show Taskpane** add-in button displayed in the ribbon.</span></span> <span data-ttu-id="85125-206">アドインは IIS 上でローカルにホストされます。</span><span class="sxs-lookup"><span data-stu-id="85125-206">The add-in will be hosted locally on IIS.</span></span>

    ![[開始] ボタンが強調表示されている Visual Studio のスクリーンショット](../images/powerpoint-tutorial-start.png)

2. <span data-ttu-id="85125-208">PowerPoint でリボンの **[作業ウィンドウの表示]** ボタンをクリックし、アドインの作業ウィンドウを開きます。</span><span class="sxs-lookup"><span data-stu-id="85125-208">In PowerPoint, select the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![[ホーム] リボンで [作業ウィンドウの表示] ボタンが強調表示されている Visual Studio のスクリーンショット](../images/powerpoint-tutorial-show-taskpane-button.png)

3. <span data-ttu-id="85125-p123">作業ウィンドウで **[Get Slide Metadata]** (スライドのメタデータの取得) ボタンを選択し、選択したスライドのメタデータを取得します。 スライドのメタデータは作業ウィンドウの下部にあるポップアップ ダイアログ ウィンドウに書き込まれます。 この例では、JSON メタデータ内の `slides` 配列に、選択したスライドの `id`、`title`、および `index` を指定するオブジェクトが 1 つ含まれます。 スライドのメタデータを取得するときに複数のスライドが選択されている場合、JSON メタデータ内の `slides` 配列には、選択したスライドごとにオブジェクトが 1 つ含まれます。</span><span class="sxs-lookup"><span data-stu-id="85125-p123">In the task pane, choose the **Get Slide Metadata** button to get the metadata for the selected slide. The slide metadata is written to the popup dialog window at the bottom of the task pane. In this case, the `slides` array within the JSON metadata contains one object that specifies the `id`, `title`, and `index` of the selected slide. If multiple slides had been selected when you retrieved slide metadata, the `slides` array within the JSON metadata would contain one object for each selected slide.</span></span>

    ![[Get Slide Metadata] (スライドのメタデータの取得) ボタンが強調表示されている PowerPoint アドインのスクリーンショット](../images/powerpoint-tutorial-get-slide-metadata.png)

4. <span data-ttu-id="85125-215">Visual Studio で **Shift + F5** を押すか **[停止]** ボタンを選択してアドインを停止します。</span><span class="sxs-lookup"><span data-stu-id="85125-215">In Visual Studio, stop the add-in by pressing **Shift + F5** or choosing the **Stop** button.</span></span> <span data-ttu-id="85125-216">アドインが停止すると、PowerPoint は自動的に閉じます。</span><span class="sxs-lookup"><span data-stu-id="85125-216">PowerPoint will automatically close when the add-in is stopped.</span></span>

    ![[停止] ボタンが強調表示されている Visual Studio のスクリーンショット](../images/powerpoint-tutorial-stop.png)

## <a name="navigate-between-slides"></a><span data-ttu-id="85125-218">スライド間の移動</span><span class="sxs-lookup"><span data-stu-id="85125-218">Navigate between slides</span></span>

<span data-ttu-id="85125-219">ドキュメントのスライド間を移動するコードを追加するには、次の手順を実行します。</span><span class="sxs-lookup"><span data-stu-id="85125-219">Complete the following steps to add code that navigates between the slides of a document.</span></span>

1. <span data-ttu-id="85125-p125">**Home.html** ファイルで `TODO5` を次のマークアップに置き換えます。 このマークアップにより、アドインの作業ウィンドウ内に表示される 4 つのナビゲーション ボタンを定義します。</span><span class="sxs-lookup"><span data-stu-id="85125-p125">In the **Home.html** file, replace `TODO5` with the following markup. This markup defines the four navigation buttons that will appear within the add-in's task pane.</span></span>

    ```html
    <br /><br />
    <button class="Button Button--primary" id="go-to-first-slide">
        <span class="Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="Button-label">Go to First Slide</span>
        <span class="Button-description">Go to the first slide.</span>
    </button>
    <br /><br />
    <button class="Button Button--primary" id="go-to-next-slide">
        <span class="Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="Button-label">Go to Next Slide</span>
        <span class="Button-description">Go to the next slide.</span>
    </button>
    <br /><br />
    <button class="Button Button--primary" id="go-to-previous-slide">
        <span class="Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="Button-label">Go to Previous Slide</span>
        <span class="Button-description">Go to the previous slide.</span>
    </button>
    <br /><br />
    <button class="Button Button--primary" id="go-to-last-slide">
        <span class="Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="Button-label">Go to Last Slide</span>
        <span class="Button-description">Go to the last slide.</span>
    </button>
    ```

2. <span data-ttu-id="85125-222">**Home.js** ファイルで `TODO8` を次のコードに置き換え、4 つのナビゲーション ボタンのイベント ハンドラーを割り当てます。</span><span class="sxs-lookup"><span data-stu-id="85125-222">In the **Home.js** file, replace `TODO8` with the following code to assign the event handlers for the four navigation buttons.</span></span>

    ```js
    $('#go-to-first-slide').click(goToFirstSlide);
    $('#go-to-next-slide').click(goToNextSlide);
    $('#go-to-previous-slide').click(goToPreviousSlide);
    $('#go-to-last-slide').click(goToLastSlide);
    ```

3. <span data-ttu-id="85125-223">**Home.js** ファイルで `TODO9` を次のコードに置き換え、ナビゲーション関数を定義します。</span><span class="sxs-lookup"><span data-stu-id="85125-223">In the **Home.js** file, replace `TODO9` with the following code to define the navigation functions.</span></span> <span data-ttu-id="85125-224">これらの関数では `goToByIdAsync` 関数を使用して、ドキュメント内のその位置 (最初、最後、前、次) に基づいてスライドを選択します。</span><span class="sxs-lookup"><span data-stu-id="85125-224">Each of these functions uses the `goToByIdAsync` function to select a slide based upon its position in the document (first, last, previous, and next).</span></span>

    ```js
    function goToFirstSlide() {
        Office.context.document.goToByIdAsync(Office.Index.First, Office.GoToType.Index,
            function (asyncResult) {
                if (asyncResult.status == "failed") {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }

    function goToLastSlide() {
        Office.context.document.goToByIdAsync(Office.Index.Last, Office.GoToType.Index,
            function (asyncResult) {
                if (asyncResult.status == "failed") {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }

    function goToPreviousSlide() {
        Office.context.document.goToByIdAsync(Office.Index.Previous, Office.GoToType.Index,
            function (asyncResult) {
                if (asyncResult.status == "failed") {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }

    function goToNextSlide() {
        Office.context.document.goToByIdAsync(Office.Index.Next, Office.GoToType.Index,
            function (asyncResult) {
                if (asyncResult.status == "failed") {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }
    ```

### <a name="test-the-add-in"></a><span data-ttu-id="85125-225">アドインをテストする</span><span class="sxs-lookup"><span data-stu-id="85125-225">Test the add-in</span></span>

1. <span data-ttu-id="85125-226">Visual Studio を使用して、アドインをテストします。そのために、**F5** キーを押すか **[開始]** ボタンをクリックして、リボンに **[作業ウィンドウの表示]** アドイン ボタンが表示された PowerPoint を起動します。</span><span class="sxs-lookup"><span data-stu-id="85125-226">Using Visual Studio, test the add-in by pressing **F5** or choosing the **Start** button to launch PowerPoint with the **Show Taskpane** add-in button displayed in the ribbon.</span></span> <span data-ttu-id="85125-227">アドインは IIS 上でローカルにホストされます。</span><span class="sxs-lookup"><span data-stu-id="85125-227">The add-in will be hosted locally on IIS.</span></span>

    ![[開始] ボタンが強調表示されている Visual Studio のスクリーンショット](../images/powerpoint-tutorial-start.png)

2. <span data-ttu-id="85125-229">PowerPoint でリボンの **[作業ウィンドウの表示]** ボタンをクリックし、アドインの作業ウィンドウを開きます。</span><span class="sxs-lookup"><span data-stu-id="85125-229">In PowerPoint, select the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![[ホーム] リボンで [作業ウィンドウの表示] ボタンが強調表示されている Visual Studio のスクリーンショット](../images/powerpoint-tutorial-show-taskpane-button.png)


3. <span data-ttu-id="85125-231">**[ホーム]** タブの **[新しいスライド]** ボタンを使用して、2 つの新しいスライドをドキュメントに追加します。</span><span class="sxs-lookup"><span data-stu-id="85125-231">Use the **New Slide** button in the ribbon of the **Home** tab to add two new slides to the document.</span></span> 

4. <span data-ttu-id="85125-p128">作業ウィンドウで **[最初のスライドに移動]** ボタンをクリックします。 ドキュメントの最初のスライドが選択され、表示されます。</span><span class="sxs-lookup"><span data-stu-id="85125-p128">In the task pane, choose the **Go to First Slide** button. The first slide in the document is selected and displayed.</span></span>

    ![[最初のスライドに移動] ボタンが強調表示されている PowerPoint アドインのスクリーンショット](../images/powerpoint-tutorial-go-to-first-slide.png)

5. <span data-ttu-id="85125-p129">作業ウィンドウで **[次のスライドに移動]** ボタンをクリックします。 ドキュメントの次のスライドが選択され、表示されます。</span><span class="sxs-lookup"><span data-stu-id="85125-p129">In the task pane, choose the **Go to Next Slide** button. The next slide in the document is selected and displayed.</span></span>

    ![[次のスライドに移動] ボタンが強調表示されている PowerPoint アドインのスクリーンショット](../images/powerpoint-tutorial-go-to-next-slide.png)

6. <span data-ttu-id="85125-p130">作業ウィンドウで **[前のスライドに移動]** ボタンをクリックします。 ドキュメントの前のスライドが選択され、表示されます。</span><span class="sxs-lookup"><span data-stu-id="85125-p130">In the task pane, choose the **Go to Previous Slide** button. The previous slide in the document is selected and displayed.</span></span>

    ![[前のスライドに移動] ボタンが強調表示されている PowerPoint アドインのスクリーンショット](../images/powerpoint-tutorial-go-to-previous-slide.png)

7. <span data-ttu-id="85125-p131">作業ウィンドウで **[最後のスライドに移動]** ボタンをクリックします。 ドキュメントの最後のスライドが選択され、表示されます。</span><span class="sxs-lookup"><span data-stu-id="85125-p131">In the task pane, choose the **Go to Last Slide** button. The last slide in the document is selected and displayed.</span></span>

    ![[最後のスライドに移動] ボタンが強調表示されている PowerPoint アドインのスクリーンショット](../images/powerpoint-tutorial-go-to-last-slide.png)

8. <span data-ttu-id="85125-244">Visual Studio で **Shift + F5** を押すか **[停止]** ボタンを選択してアドインを停止します。</span><span class="sxs-lookup"><span data-stu-id="85125-244">In Visual Studio, stop the add-in by pressing **Shift + F5** or choosing the **Stop** button.</span></span> <span data-ttu-id="85125-245">アドインが停止すると、PowerPoint は自動的に閉じます。</span><span class="sxs-lookup"><span data-stu-id="85125-245">PowerPoint will automatically close when the add-in is stopped.</span></span>

    ![[停止] ボタンが強調表示されている Visual Studio のスクリーンショット](../images/powerpoint-tutorial-stop.png)

## <a name="next-steps"></a><span data-ttu-id="85125-247">次の手順</span><span class="sxs-lookup"><span data-stu-id="85125-247">Next steps</span></span>

<span data-ttu-id="85125-248">このチュートリアルでは、画像の挿入、テキストの挿入、スライドのメタデータ取得、およびスライド間の移動のための PowerPoint アドインを作成しました。</span><span class="sxs-lookup"><span data-stu-id="85125-248">In this tutorial, you've created a PowerPoint add-in that inserts an image, inserts text, gets slide metadata, and navigates between slides.</span></span> <span data-ttu-id="85125-249">PowerPoint アドインの構築に関する詳細については、次の記事にお進みください。</span><span class="sxs-lookup"><span data-stu-id="85125-249">To learn more about building PowerPoint add-ins, continue to the following article:</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="85125-250">PowerPoint アドインの概要</span><span class="sxs-lookup"><span data-stu-id="85125-250">PowerPoint add-ins overview</span></span>](../powerpoint/powerpoint-add-ins.md)
