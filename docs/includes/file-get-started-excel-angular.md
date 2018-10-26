# <a name="build-an-excel-add-in-using-angular"></a><span data-ttu-id="676b9-101">Angular を使用して Excel のアドインを作成する</span><span class="sxs-lookup"><span data-stu-id="676b9-101">Build an Excel add-in using Angular</span></span>

<span data-ttu-id="676b9-102">この記事では、Angular と Excel の JavaScript API を使用して Excel アドインを構築する手順について説明します。</span><span class="sxs-lookup"><span data-stu-id="676b9-102">In this article, you'll walk you through the process of building an Excel add-in using Angular and the Excel JavaScript API.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="676b9-103">前提条件</span><span class="sxs-lookup"><span data-stu-id="676b9-103">Prerequisites</span></span>

- [<span data-ttu-id="676b9-104">Node.js</span><span class="sxs-lookup"><span data-stu-id="676b9-104">Node.js</span></span>](https://nodejs.org)

- <span data-ttu-id="676b9-105">[Yeoman](https://github.com/yeoman/yo) の最新バージョンと [Office アドイン用の Yeoman ジェネレーター](https://github.com/OfficeDev/generator-office)をグローバルにインストールします。</span><span class="sxs-lookup"><span data-stu-id="676b9-105">Install the latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globally.</span></span>

    ```bash
    npm install -g yo generator-office
    ```

## <a name="create-the-web-app"></a><span data-ttu-id="676b9-106">Web アプリを作成する</span><span class="sxs-lookup"><span data-stu-id="676b9-106">Create the web app</span></span>

1. <span data-ttu-id="676b9-107">Yeoman のジェネレーターを使用して、Excel アドイン プロジェクトを作成します。</span><span class="sxs-lookup"><span data-stu-id="676b9-107">Use the Yeoman generator to create an Outlook add-in project.</span></span> <span data-ttu-id="676b9-108">次のコマンドを実行し、以下のプロンプトに応答します。</span><span class="sxs-lookup"><span data-stu-id="676b9-108">Run the following command and then answer the prompts as follows:</span></span>

    ```bash
    yo office
    ```

    - <span data-ttu-id="676b9-109">**プロジェクトタイプを選択してください** `Office Add-in project using Angular framework`</span><span class="sxs-lookup"><span data-stu-id="676b9-109">**Choose a project type:** `Office Add-in project using Angular framework`</span></span>
    - <span data-ttu-id="676b9-110">**Choose a script type: (スクリプト タイプを選択してください)** `Typescript`</span><span class="sxs-lookup"><span data-stu-id="676b9-110">**Choose a script type:** `Typescript`</span></span>
    - <span data-ttu-id="676b9-111">**What would you want to name your add-in?: (アドインの名前を何にしますか)** `My Office Add-in`</span><span class="sxs-lookup"><span data-stu-id="676b9-111">**What do you want to name your add-in?:** `My Office Add-in`</span></span>
    - <span data-ttu-id="676b9-112">**Which Office client application would you like to support?: (どの Office クライアント アプリケーションをサポートしますか)** `Excel`</span><span class="sxs-lookup"><span data-stu-id="676b9-112">**Which Office client application would you like to support?:** `Excel`</span></span>

    ![Yeoman ジェネレーター](../images/yo-office-excel-angular.png)
    
    <span data-ttu-id="676b9-114">ウィザードが完了すると、ジェネレーターはプロジェクトを作成し、サポートする Node コンポーネントをインストールします。</span><span class="sxs-lookup"><span data-stu-id="676b9-114">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>

2. <span data-ttu-id="676b9-115">プロジェクトのルート フォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="676b9-115">Navigate to the root folder of the web application project.</span></span>

    ```bash
    cd "My Office Add-in"
    ```

## <a name="update-the-code"></a><span data-ttu-id="676b9-116">コードを更新する</span><span class="sxs-lookup"><span data-stu-id="676b9-116">Update the code</span></span>

1. <span data-ttu-id="676b9-117">コード エディターで、　**app.css** のファイルを開き、ファイルの末尾に次のスタイルを追加し、ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="676b9-117">In your code editor, open the file **app.css**, add the following styles to the end of the file, and save the file.</span></span>

    ```css
    #content-header {
        background: #2a8dd4;
        color: #fff;
        position: absolute;
        top: 0;
        left: 0;
        width: 100%;
        height: 80px; 
        overflow: hidden;
        font-family: Arial;
        padding-top: 25px;
    }

    #content-main {
        background: #fff;
        position: fixed;
        top: 80px;
        left: 0;
        right: 0;
        bottom: 0;
        overflow: auto; 
        font-family: Arial;
    }

    .padding {
        padding: 15px;
    }

    .padding-sm {
        padding: 4px;
    }

    .normal-button {
        width: 80px;
        padding: 2px;
    }
    ```

2. <span data-ttu-id="676b9-118"> *\*src/app/app.component.html** のファイルを開き、次のコードで全体のコンテンツを置き換え、そのファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="676b9-118">Open **src/app/app.component.html**, replace file contents with the following HTML, and save the file.</span></span>

    ```html
    <div id="content-header">
        <div class="padding">
            <h1>{{welcomeMessage}}</h1>
        </div>
    </div>
    <div id="content-main">
        <div class="padding">
            <p>Choose the button below to set the color of the selected range to green.</p>
            <br />
            <h3>Try it out</h3>
            <br />
            <div role="button" class="ms-Button" (click)="setColor()">
                <span class="ms-Button-label">Set color</span>
                <span class="ms-Button-icon"><i class="ms-Icon ms-Icon--ChevronRight"></i></span>
            </div>
        </div>
    </div>
    ```

3. <span data-ttu-id="676b9-119"> *\*src/app/app.component.ts** のファイルを開き、次のコードで全体のコンテンツを置き換え、そのファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="676b9-119">Open **src/app/app.component.ts**, replace file contents with the following code, and save the file.</span></span>

    ```typescript
    import { Component } from '@angular/core';
    import * as OfficeHelpers from '@microsoft/office-js-helpers';

    const template = require('./app.component.html');

    @Component({
        selector: 'app-home',
        template
    })
    export default class AppComponent {
        welcomeMessage = 'Welcome';

        async setColor() {
            try {
                await Excel.run(async context => {
                    const range = context.workbook.getSelectedRange();
                    range.load('address');
                    range.format.fill.color = 'green';
                    await context.sync();
                    console.log(`The range address was ${range.address}.`);
                });
            } catch (error) {
                OfficeHelpers.UI.notify(error);
                OfficeHelpers.Utilities.log(error);
            }
        }

    }
    ```

## <a name="update-the-manifest"></a><span data-ttu-id="676b9-120">マニフェストを更新する</span><span class="sxs-lookup"><span data-stu-id="676b9-120">Update the manifest</span></span>

1. <span data-ttu-id="676b9-121"> *\*manifest.xml** のファイルを開き、アドインの設定と機能を定義します。</span><span class="sxs-lookup"><span data-stu-id="676b9-121">Open the file **one-note-add-in-manifest.xml** to define the add-in's settings and capabilities.</span></span> 

2. <span data-ttu-id="676b9-122"> `ProviderName` 要素にはプレースホルダー値が含まれています。</span><span class="sxs-lookup"><span data-stu-id="676b9-122">The `ProviderName` element has a placeholder value.</span></span> <span data-ttu-id="676b9-123">それを自分の名前に置き換えます。</span><span class="sxs-lookup"><span data-stu-id="676b9-123">Replace it with your name.</span></span>

3. <span data-ttu-id="676b9-p103"> `Description` 要素の `DefaultValue` 属性にはプレースホルダーがあります。これを *\*Excel の作業ウィンドウ アドイン** に置き換えます。</span><span class="sxs-lookup"><span data-stu-id="676b9-p103">The `DefaultValue` attribute of the `Description` element has a placeholder. Replace it with **A task pane add-in for Excel**.</span></span>

4. <span data-ttu-id="676b9-126">ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="676b9-126">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Excel"/>
    ...
    ```

## <a name="start-the-dev-server"></a><span data-ttu-id="676b9-127">開発用サーバーを起動する</span><span class="sxs-lookup"><span data-stu-id="676b9-127">Start the dev server</span></span>

[!include[Start server section](../includes/quickstart-yo-start-server.md)] 

## <a name="try-it-out"></a><span data-ttu-id="676b9-128">試してみる</span><span class="sxs-lookup"><span data-stu-id="676b9-128">Try it out</span></span>

1. <span data-ttu-id="676b9-129">アドインを実行して、Excel 内のアドインをサイドロードするために使用するプラットフォームの手順に従います。</span><span class="sxs-lookup"><span data-stu-id="676b9-129">Follow the instructions for the platform you'll be using to run your add-in and sideload the add-in within Excel.</span></span>

    - <span data-ttu-id="676b9-130">Windows: [Windows で Office アドインをサイドロードする](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="676b9-130">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
    - <span data-ttu-id="676b9-131">Excel Online:[Office Online で Office アドインをサイドロードする](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span><span class="sxs-lookup"><span data-stu-id="676b9-131">Excel Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span></span>
    - <span data-ttu-id="676b9-132">iPad および Mac: [iPad と Mac で Office アドインをサイドロードする](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="676b9-132">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

   
2. <span data-ttu-id="676b9-133">Excel で、**[ホーム]** タブを選択し、リボンの **[作業ウィンドウの表示]** ボタンをクリックして、アドインの作業ウィンドウを開きます。</span><span class="sxs-lookup"><span data-stu-id="676b9-133">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Excel アドイン ボタン](../images/excel-quickstart-addin-2b.png)

3. <span data-ttu-id="676b9-135">ワークシート内で任意のセル範囲を選択します。</span><span class="sxs-lookup"><span data-stu-id="676b9-135">Select any range of cells in the worksheet.</span></span>

4. <span data-ttu-id="676b9-136">作業ウィンドウで、**[色の設定]** ボタンをクリックして、選択範囲の色を緑に設定します。</span><span class="sxs-lookup"><span data-stu-id="676b9-136">In the task pane, choose the **Set color** button to set the color of the selected range to green.</span></span>

    ![Excel アドイン](../images/excel-quickstart-addin-2c.png)

## <a name="next-steps"></a><span data-ttu-id="676b9-138">次の手順</span><span class="sxs-lookup"><span data-stu-id="676b9-138">Next steps</span></span>

<span data-ttu-id="676b9-p104">これで完了です。Angular を使用して Excel アドインが正常に作成されました。次に、Excel アドインの機能の詳細について説明します。Excel アドインのチュートリアルに従って、より複雑なアドインをビルドします。</span><span class="sxs-lookup"><span data-stu-id="676b9-p104">Congratulations, you've successfully created an Excel add-in using Angular! Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="676b9-141">Excel アドインのチュートリアル</span><span class="sxs-lookup"><span data-stu-id="676b9-141">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.yml)

## <a name="see-also"></a><span data-ttu-id="676b9-142">関連項目</span><span class="sxs-lookup"><span data-stu-id="676b9-142">See also</span></span>

* [<span data-ttu-id="676b9-143">Excel アドインのチュートリアル</span><span class="sxs-lookup"><span data-stu-id="676b9-143">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial-create-table.md)
* [<span data-ttu-id="676b9-144">Excel JavaScript API を使用した基本的なプログラミングの概念</span><span class="sxs-lookup"><span data-stu-id="676b9-144">Fundamental programming concepts with the Excel JavaScript API</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="676b9-145">Excel アドインのコード サンプル</span><span class="sxs-lookup"><span data-stu-id="676b9-145">Excel add-in code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="676b9-146">Excel JavaScript API リファレンス</span><span class="sxs-lookup"><span data-stu-id="676b9-146">Excel JavaScript API reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview?view=office-js)
