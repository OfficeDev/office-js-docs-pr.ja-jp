# <a name="build-an-excel-add-in-using-react"></a><span data-ttu-id="24066-101">React を使用して Excel のアドインを作成する</span><span class="sxs-lookup"><span data-stu-id="24066-101">Build an Excel add-in using React</span></span>

<span data-ttu-id="24066-102">この記事では、React と Excel の JavaScript API を使用して Excel アドインを構築する手順について説明します。</span><span class="sxs-lookup"><span data-stu-id="24066-102">In this article, you'll walk through the process of building an Excel add-in using React and the Excel JavaScript API.</span></span>

## <a name="environment"></a><span data-ttu-id="24066-103">環境</span><span class="sxs-lookup"><span data-stu-id="24066-103">Environment</span></span>

- <span data-ttu-id="24066-104">**Office Desktop**最新バージョンのOfficeがインストールされていることを確認してください。</span><span class="sxs-lookup"><span data-stu-id="24066-104">**Office Desktop**: Ensure that you have the latest version of Office installed.</span></span> <span data-ttu-id="24066-105">アドインコマンドにはビルド16.0.6769.0000以上が必要です（**16.0.6868.0000** 推奨）。</span><span class="sxs-lookup"><span data-stu-id="24066-105">Add-in commands require build 16.0.6769.0000 or higher (**16.0.6868.0000** recommended).</span></span> <span data-ttu-id="24066-106">[Officeアプリケーションの最新バージョンをインストールする](http://aka.ms/latestoffice)のやり方を学ぼう。</span><span class="sxs-lookup"><span data-stu-id="24066-106">Learn how to [Install the latest version of Office applications](http://aka.ms/latestoffice).</span></span> 
 
- <span data-ttu-id="24066-107">**Office Online**：追加設定はありません。</span><span class="sxs-lookup"><span data-stu-id="24066-107">**Office Online**: There is no additional setup.</span></span> <span data-ttu-id="24066-108">Office Online の職場/学校アカウント用コマンドのサポートはプレビューになっています。</span><span class="sxs-lookup"><span data-stu-id="24066-108">Please note that support for commands in Office Online for work/school accounts is in preview.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="24066-109">前提条件</span><span class="sxs-lookup"><span data-stu-id="24066-109">Prerequisites</span></span>

- <span data-ttu-id="24066-110">[Create React App](https://github.com/facebookincubator/create-react-app) をグローバルにインストールします。</span><span class="sxs-lookup"><span data-stu-id="24066-110">Install [Create React App](https://github.com/facebookincubator/create-react-app) globally.</span></span>

    ```bash
    npm install -g create-react-app
    ```

- <span data-ttu-id="24066-111">[Yeoman](https://github.com/yeoman/yo) の最新バージョンと [Office アドイン用の Yeoman ジェネレーター](https://github.com/OfficeDev/generator-office)をグローバルにインストールします。</span><span class="sxs-lookup"><span data-stu-id="24066-111">Install the latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globally.</span></span>

    ```bash
    npm install -g yo generator-office
    ```

## <a name="generate-a-new-react-app"></a><span data-ttu-id="24066-112">新しい React アプリを生成する</span><span class="sxs-lookup"><span data-stu-id="24066-112">Generate a new React app</span></span>

<span data-ttu-id="24066-113">Create React App を使用して、React アプリを生成します。</span><span class="sxs-lookup"><span data-stu-id="24066-113">Use Create React App to generate your React app.</span></span> <span data-ttu-id="24066-114">ターミナルから、次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="24066-114">From the terminal, run the following command:</span></span>

```bash
create-react-app my-addin
```

## <a name="generate-the-manifest-file-and-sideload-the-add-in"></a><span data-ttu-id="24066-115">マニフェスト ファイルを生成し、アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="24066-115">Generate the manifest file and sideload the add-in</span></span>

<span data-ttu-id="24066-116">各アドインには、設定と機能を定義するマニフェスト ファイルが必要です。</span><span class="sxs-lookup"><span data-stu-id="24066-116">Each add-in requires a manifest file to define its settings and capabilities.</span></span>

1. <span data-ttu-id="24066-117">アプリ フォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="24066-117">Navigate to your app folder.</span></span>

    ```bash
    cd my-addin
    ```

2. <span data-ttu-id="24066-118">Yeoman ジェネレーター使用して、アドインのマニフェスト ファイルを生成します。</span><span class="sxs-lookup"><span data-stu-id="24066-118">Use the Yeoman generator to generate the manifest file for your add-in.</span></span> <span data-ttu-id="24066-119">次のコマンドを実行し、以下のスクリーンショットに示すとおり、プロンプトに応答します。</span><span class="sxs-lookup"><span data-stu-id="24066-119">Run the following command and then answer the prompts as shown in the following screenshot:</span></span>

    ```bash
    yo office 
    ```

    - <span data-ttu-id="24066-120">**Choose a project type:​ (プロジェクト タイプを選択してください)** `Office Add-in containing the manifest only`</span><span class="sxs-lookup"><span data-stu-id="24066-120">**Choose a project type:** `Office Add-in containing the manifest only`</span></span>
    - <span data-ttu-id="24066-121">**What would you want to name your add-in?: (アドインの名前を何にしますか)** `My Office Add-in`</span><span class="sxs-lookup"><span data-stu-id="24066-121">**What do you want to name your add-in?:** `My Office Add-in`</span></span>
    - <span data-ttu-id="24066-122">**Which Office client application would you like to support?: (どの Office クライアント アプリケーションをサポートしますか)** `Excel`</span><span class="sxs-lookup"><span data-stu-id="24066-122">**Which Office client application would you like to support?:** `Excel`</span></span>


    <span data-ttu-id="24066-123">ウィザードを完了すると、マニフェスト ファイルとリソース ファイルを使用してプロジェクトをビルドできます。</span><span class="sxs-lookup"><span data-stu-id="24066-123">After you complete the wizard, a manifest file and resource file are available for you to build your project.</span></span>
    
    ![Yeoman ジェネレーター](../images/yo-office.png)
    
    > [!NOTE]
    > <span data-ttu-id="24066-125">**package.json** を上書きするメッセージが表示された場合は、**No** (上書きしない) と応答します。</span><span class="sxs-lookup"><span data-stu-id="24066-125">If you're prompted to overwrite **package.json**, answer **No** (do not overwrite).</span></span>

3. <span data-ttu-id="24066-126">アドインを実行して、Excel 内のアドインをサイドロードするのに使用するプラットフォームの手順に従います。</span><span class="sxs-lookup"><span data-stu-id="24066-126">Follow the instructions for the platform you'll be using to run your add-in and sideload the add-in within Excel.</span></span>

    - <span data-ttu-id="24066-127">Windows: [Windows で Office アドインをサイドロードする](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="24066-127">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
    - <span data-ttu-id="24066-128">Excel Online:[Office Online で Office アドインをサイドロードする](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span><span class="sxs-lookup"><span data-stu-id="24066-128">Excel Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span></span>
    - <span data-ttu-id="24066-129">iPad および Mac:[iPad と Mac で Office アドインをサイドロードする](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="24066-129">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

## <a name="update-the-app"></a><span data-ttu-id="24066-130">アプリを更新する</span><span class="sxs-lookup"><span data-stu-id="24066-130">Update the app</span></span>

1. <span data-ttu-id="24066-131">**public/index.html** を開き、`</head>` タグの直前に次の `<script>` タグを追加し、ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="24066-131">Open **public/index.html**, add the following `<script>` tag immediately before the `</head>` tag, and save the file.</span></span>

    ```html
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
    ```

2. <span data-ttu-id="24066-132">**src/index.js** を開き、`ReactDOM.render(<App />, document.getElementById('root'));` を次のコードで置き換えて、ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="24066-132">Open **src/index.js**, replace `ReactDOM.render(<App />, document.getElementById('root'));` with the following code, and save the file.</span></span> 

    ```typescript
    const Office = window.Office;
    
    Office.initialize = () => {
      ReactDOM.render(<App />, document.getElementById('root'));
    };
    ```

3. <span data-ttu-id="24066-133">**src/App.js** を開き、ファイルのコンテンツを次のコードで置き換えて、ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="24066-133">Open **src/App.js**, replace file contents with the following code, and save the file.</span></span> 

    ```js
    import React, { Component } from 'react';
    import './App.css';

    class App extends Component {
      constructor(props) {
        super(props);

        this.onSetColor = this.onSetColor.bind(this);
      }

      onSetColor() {
        window.Excel.run(async (context) => {
          const range = context.workbook.getSelectedRange();
          range.format.fill.color = 'green';
          await context.sync();
        });
      }

      render() {
        return (
          <div id="content">
            <div id="content-header">
              <div className="padding">
                  <h1>Welcome</h1>
              </div>
            </div>
            <div id="content-main">
              <div className="padding">
                  <p>Choose the button below to set the color of the selected range to green.</p>
                  <br />
                  <h3>Try it out</h3>
                  <button onClick={this.onSetColor}>Set color</button>
              </div>
            </div>
          </div>
        );
      }
    }

    export default App;
    ```

4. <span data-ttu-id="24066-134">**src/App.css** を開き、ファイルのコンテンツを次の CSS コードで置き換えて、ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="24066-134">Open **src/App.css**, replace file contents with the following CSS code, and save the file.</span></span> 

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
    }

    #content-main {
        background: #fff;
        position: fixed;
        top: 80px;
        left: 0;
        right: 0;
        bottom: 0;
        overflow: auto; 
    }

    .padding {
        padding: 15px;
    }
    ```

## <a name="try-it-out"></a><span data-ttu-id="24066-135">お試しください</span><span class="sxs-lookup"><span data-stu-id="24066-135">Try it out</span></span>

1. <span data-ttu-id="24066-136">ターミナルから、次のコマンドを実行してデベロッパー サーバーを起動します。</span><span class="sxs-lookup"><span data-stu-id="24066-136">From the terminal, run the following command to start the dev server.</span></span>

    <span data-ttu-id="24066-137">Windows</span><span class="sxs-lookup"><span data-stu-id="24066-137">Windows:</span></span>
    ```bash
    set HTTPS=true&&npm start
    ```

    <span data-ttu-id="24066-138">macOS</span><span class="sxs-lookup"><span data-stu-id="24066-138">macOS:</span></span>
    ```bash
    HTTPS=true npm start
    ```

   > [!NOTE]
   > <span data-ttu-id="24066-p105">アドインが含まれているブラウザー ウィンドウが開きます。このウィンドウを閉じます。</span><span class="sxs-lookup"><span data-stu-id="24066-p105">A browser window will open with the add-in in it. Close this window.</span></span>

2. <span data-ttu-id="24066-141">Excel で、**[ホーム]** タブを選択し、リボンの **[作業ウィンドウの表示]** ボタンをクリックして、アドインの作業ウィンドウを開きます。</span><span class="sxs-lookup"><span data-stu-id="24066-141">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Excel アドイン ボタン](../images/excel-quickstart-addin-2b.png)

3. <span data-ttu-id="24066-143">ワークシート内で任意のセルの範囲を選択します。</span><span class="sxs-lookup"><span data-stu-id="24066-143">Select any range of cells in the worksheet.</span></span>

4. <span data-ttu-id="24066-144">作業ウィンドウで、**[色の設定]** ボタンをクリックして、選択範囲の色を緑に設定します。</span><span class="sxs-lookup"><span data-stu-id="24066-144">In the task pane, choose the **Set color** button to set the color of the selected range to green.</span></span>

    ![Excel アドイン](../images/excel-quickstart-addin-2c.png)

## <a name="next-steps"></a><span data-ttu-id="24066-146">次の手順</span><span class="sxs-lookup"><span data-stu-id="24066-146">Next steps</span></span>

<span data-ttu-id="24066-p106">これで完了です。React を使用して Excel アドインが正常に作成されました。次に、Excel アドインの機能の詳細について説明します。Excel アドインのチュートリアルに従って、より複雑なアドインをビルドします。</span><span class="sxs-lookup"><span data-stu-id="24066-p106">Congratulations, you've successfully created an Excel add-in using React! Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="24066-149">Excel アドインのチュートリアル</span><span class="sxs-lookup"><span data-stu-id="24066-149">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.yml)

## <a name="see-also"></a><span data-ttu-id="24066-150">関連項目</span><span class="sxs-lookup"><span data-stu-id="24066-150">See also</span></span>

* [<span data-ttu-id="24066-151">Excel アドインのチュートリアル</span><span class="sxs-lookup"><span data-stu-id="24066-151">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial-create-table.md)
* [<span data-ttu-id="24066-152">Excel JavaScript API の中心概念</span><span class="sxs-lookup"><span data-stu-id="24066-152">Excel JavaScript API core concepts</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="24066-153">Excel アドインのコード サンプル</span><span class="sxs-lookup"><span data-stu-id="24066-153">Excel add-in code samples</span></span>](http://dev.office.com/code-samples#?filters=excel,office%20add-ins)
* [<span data-ttu-id="24066-154">Excel JavaScript API リファレンス</span><span class="sxs-lookup"><span data-stu-id="24066-154">Excel JavaScript API reference</span></span>](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview)
