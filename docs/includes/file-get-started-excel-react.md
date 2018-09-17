# <a name="build-an-excel-add-in-using-react"></a><span data-ttu-id="b7040-101">React を使用して Excel のアドインを作成する</span><span class="sxs-lookup"><span data-stu-id="b7040-101">Build an Excel add-in using React</span></span>

<span data-ttu-id="b7040-102">この記事では、React と Excel の JavaScript API を使用して Excel アドインを構築する手順について説明します。</span><span class="sxs-lookup"><span data-stu-id="b7040-102">In this article, you'll walk through the process of building an Excel add-in using React and the Excel JavaScript API.</span></span>

## <a name="environment"></a><span data-ttu-id="b7040-103">環境</span><span class="sxs-lookup"><span data-stu-id="b7040-103">Environment</span></span>

- <span data-ttu-id="b7040-104">**Office Desktop**最新バージョンのOfficeがインストールされていることを確認してください。</span><span class="sxs-lookup"><span data-stu-id="b7040-104">**Office Desktop**: Ensure that you have the latest version of Office installed.</span></span> <span data-ttu-id="b7040-105">アドインコマンドにはビルド 16.0.6769.0000 以上が必要です (**16.0.6868.0000** 推奨)。</span><span class="sxs-lookup"><span data-stu-id="b7040-105">Add-in commands require build 16.0.6769.0000 or higher (**16.0.6868.0000** recommended).</span></span> <span data-ttu-id="b7040-106">[Office アプリケーションの最新バージョンをインストールする](http://aka.ms/latestoffice)方法。</span><span class="sxs-lookup"><span data-stu-id="b7040-106">Learn how to [Install the latest version of Office applications](http://aka.ms/latestoffice).</span></span> 
 
- <span data-ttu-id="b7040-107">**Office Online**：追加設定はありません。</span><span class="sxs-lookup"><span data-stu-id="b7040-107">**Office Online**: There is no additional setup.</span></span> <span data-ttu-id="b7040-108">Office Online の職場/学校アカウント用コマンドのサポートはプレビューになっています。</span><span class="sxs-lookup"><span data-stu-id="b7040-108">Please note that support for commands in Office Online for work/school accounts is in preview.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="b7040-109">前提条件</span><span class="sxs-lookup"><span data-stu-id="b7040-109">Prerequisites</span></span>

- [<span data-ttu-id="b7040-110">Node.js</span><span class="sxs-lookup"><span data-stu-id="b7040-110">Node.js</span></span>](https://nodejs.org)

- <span data-ttu-id="b7040-111">[Yeoman](https://github.com/yeoman/yo) の最新バージョンと [Office アドイン用の Yeoman ジェネレーター](https://github.com/OfficeDev/generator-office)をグローバルにインストールします。</span><span class="sxs-lookup"><span data-stu-id="b7040-111">Install the latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globally.</span></span>
    ```bash
    npm install -g yo generator-office
    ```

### <a name="create-the-web-app"></a><span data-ttu-id="b7040-112">Web アプリを作成する</span><span class="sxs-lookup"><span data-stu-id="b7040-112">Create the web app</span></span>

1. <span data-ttu-id="b7040-113">ローカル ドライブにフォルダーを作成し、**my-addin** という名前を付けます。</span><span class="sxs-lookup"><span data-stu-id="b7040-113">Create a folder on your local drive and name it **my-addin**.</span></span> <span data-ttu-id="b7040-114">ここにアプリのファイルを作成します。</span><span class="sxs-lookup"><span data-stu-id="b7040-114">This is where you'll create the files for your app.</span></span>

2. <span data-ttu-id="b7040-115">アプリ フォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="b7040-115">Navigate to your app folder.</span></span>

    ```bash
    cd my-addin
    ```

3. <span data-ttu-id="b7040-116">Yeoman ジェネレーター使用して、アドインのマニフェスト ファイルを生成します。</span><span class="sxs-lookup"><span data-stu-id="b7040-116">Use the Yeoman generator to generate the manifest file for your add-in.</span></span> <span data-ttu-id="b7040-117">次のコマンドを実行し、以下のスクリーンショットに示すとおり、プロンプトに応答します。</span><span class="sxs-lookup"><span data-stu-id="b7040-117">Run the following command and then answer the prompts as shown in the following screenshot:</span></span>

    ```bash
    yo office
    ```

    - <span data-ttu-id="b7040-118">**Choose a project type:​ (プロジェクト タイプを選択してください)** `Office Add-in project using React framework`</span><span class="sxs-lookup"><span data-stu-id="b7040-118">**Choose a project type:** `Office Add-in project using React framework`</span></span>
    - <span data-ttu-id="b7040-119">**What would you want to name your add-in?: (アドインの名前を何にしますか)** `My Office Add-in`</span><span class="sxs-lookup"><span data-stu-id="b7040-119">**What do you want to name your add-in?:** `My Office Add-in`</span></span>
    - <span data-ttu-id="b7040-120">**Which Office client application would you like to support?: (どの Office クライアント アプリケーションをサポートしますか)** `Excel`</span><span class="sxs-lookup"><span data-stu-id="b7040-120">**Which Office client application would you like to support?:** `Excel`</span></span>

    ![Yeoman ジェネレーター](../images/yo-office-excel-react.png)
    
    <span data-ttu-id="b7040-122">ウィザードが完了すると、ジェネレーターはプロジェクトを作成し、サポートする Node コンポーネントをインストールします。</span><span class="sxs-lookup"><span data-stu-id="b7040-122">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>

4.  <span data-ttu-id="b7040-123">**Src/components/App.tsx** を開き、コメントの「塗りつぶしの色を更新する」を検索し、塗りつぶしの色を '青' から'黄' に変更してから保存します。</span><span class="sxs-lookup"><span data-stu-id="b7040-123">Open **src/components/App.tsx**, search for the comment "Update the fill color," then change the fill color from 'yellow' to 'blue', and save the file.</span></span> 

    ```js
    range.format.fill.color = 'blue'

    ```

5. <span data-ttu-id="b7040-124">**src/components/App.tsx** 内の `render` 関数の `return` ブロックで、`<Herolist>` を次のコードに更新し、ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="b7040-124">In the `return` block of the `render` function within **src/components/App.tsx**, update the `<Herolist>` to the code below, and save the file.</span></span> 

    ```js
      <HeroList message='Discover what My Office Add-in can do for you today!' items={this.state.listItems}>
        <p className='ms-font-l'>Choose the button below to set the color of the selected range to blue. <b>Set color</b>.</p>
        <Button className='ms-welcome__action' buttonType={ButtonType.hero} iconProps={{ iconName: 'ChevronRight' }} onClick={this.click}>Run</Button>
    </HeroList>
    ```

6. <span data-ttu-id="b7040-125">「[自己署名証明書を信頼されたルート証明書として追加する](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md)」の手順を実行して、開発用コンピューターのオペレーティング システムの証明書を信頼します。</span><span class="sxs-lookup"><span data-stu-id="b7040-125">Carry out the steps in [Adding Self-Signed Certificates as Trusted Root Certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) to trust the certificate for your development computer's operating system.</span></span>

7. <span data-ttu-id="b7040-126">アドインをサイドロードすると Excel に表示されます。</span><span class="sxs-lookup"><span data-stu-id="b7040-126">Sideload your add-in so it will appear in Excel.</span></span> <span data-ttu-id="b7040-127">ターミナルで次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="b7040-127">In the terminal run the following command:</span></span> 
    
    ```bash
    npm run sideload
    ```

## <a name="try-it-out"></a><span data-ttu-id="b7040-128">お試しください</span><span class="sxs-lookup"><span data-stu-id="b7040-128">Try it out</span></span>

1. <span data-ttu-id="b7040-129">ターミナルから、次のコマンドを実行してデベロッパー サーバーを起動します。</span><span class="sxs-lookup"><span data-stu-id="b7040-129">From the terminal, run the following command to start the dev server.</span></span>

    <span data-ttu-id="b7040-130">Windows</span><span class="sxs-lookup"><span data-stu-id="b7040-130">Windows:</span></span>
    ```bash
    npm start
    ```

2. <span data-ttu-id="b7040-131">Excel で、**[ホーム]** タブを選択し、リボンの **[作業ウィンドウの表示]** ボタンをクリックして、アドインの作業ウィンドウを開きます。</span><span class="sxs-lookup"><span data-stu-id="b7040-131">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Excel アドイン ボタン](../images/excel-quickstart-addin-2b.png)

3. <span data-ttu-id="b7040-133">ワークシート内で任意のセルの範囲を選択します。</span><span class="sxs-lookup"><span data-stu-id="b7040-133">Select any range of cells in the worksheet.</span></span>

4. <span data-ttu-id="b7040-134">作業ウィンドウで **[色の設定]** ボタンをクリックし、選択範囲の色を青に設定します。</span><span class="sxs-lookup"><span data-stu-id="b7040-134">In the task pane, choose the **Set color** button to set the color of the selected range to green.</span></span>

    ![Excel アドイン](../images/excel-quickstart-addin-2c.png)

## <a name="next-steps"></a><span data-ttu-id="b7040-136">次の手順</span><span class="sxs-lookup"><span data-stu-id="b7040-136">Next steps</span></span>

<span data-ttu-id="b7040-p106">これで完了です。React を使用して Excel アドインが正常に作成されました。次に、Excel アドインの機能の詳細について説明します。Excel アドインのチュートリアルに従って、より複雑なアドインをビルドします。</span><span class="sxs-lookup"><span data-stu-id="b7040-p106">Congratulations, you've successfully created an Excel add-in using React! Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="b7040-139">Excel アドインのチュートリアル</span><span class="sxs-lookup"><span data-stu-id="b7040-139">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.yml)

## <a name="see-also"></a><span data-ttu-id="b7040-140">関連項目</span><span class="sxs-lookup"><span data-stu-id="b7040-140">See also</span></span>

* [<span data-ttu-id="b7040-141">Excel アドインのチュートリアル</span><span class="sxs-lookup"><span data-stu-id="b7040-141">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial-create-table.md)
* [<span data-ttu-id="b7040-142">Excel JavaScript API の中心概念</span><span class="sxs-lookup"><span data-stu-id="b7040-142">Excel JavaScript API core concepts</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="b7040-143">Excel アドインのコード サンプル</span><span class="sxs-lookup"><span data-stu-id="b7040-143">Excel add-in code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="b7040-144">Excel JavaScript API リファレンス</span><span class="sxs-lookup"><span data-stu-id="b7040-144">Excel JavaScript API reference</span></span>](https://docs.microsoft.com/javascript/office/overview/excel-add-ins-reference-overview?view=office-js)
