# <a name="build-your-first-project-add-in"></a>最初の Project アドインをビルドする

この記事では、jQuery および Office JavaScript API を使用して Project アドインをビルドする手順について説明します。

## <a name="prerequisites"></a>前提条件

- [Node.js](https://nodejs.org)

-  [Yeoman](https://github.com/yeoman/yo) および [Office アドイン用の Yeoman ジェネレーター](https://github.com/OfficeDev/generator-office) の最新バージョンをグローバルにインストールします。

    ```bash
    npm install -g yo generator-office
    ```

## <a name="create-the-add-in"></a>アドインを作成する

1. ローカルドライブにフォルダーを作成し、 `my-project-addin` という名前を付けます。ここにアドインのファイルを作成します。

    ```bash
    mkdir my-project-addin
    ```

2. 新しいフォルダーに移動します。

    ```bash
    cd my-project-addin
    ```

3. Yeomange ジェネレータを使用して Project アドインのプロジェクトを作成します。次のコマンドを実行してから、プロンプトに次のように応答します。

    ```bash
    yo office
    ```

    - **プロジェクトの型を選択します** `Office Add-in project using Jquery framework`
    - **Choose a script type: (スクリプト タイプを選択してください)** `Javascript`
    - **What would you want to name your add-in?: (アドインの名前を何にしますか)** `My Office Add-in`
    - **Which Office client application would you like to support? (サポートする Office クライアント アプリケーションを選んでください):** `Project`

    ![Yeoman ジェネレーターのプロンプトと応答のスクリーンショット](../images/yo-office-project-jquery.png)
    
    ウィザードが完了すると、ジェネレーターはプロジェクトを作成し、サポートする Node コンポーネントをインストールします。
    
4. Web アプリケーション プロジェクトのルート フォルダーに移動します。

    ```bash
    cd "My Office Add-in"
    ```

## <a name="update-the-code"></a>コードを更新する

1. コード エディターにあるプロジェクトのルートの **index.html** を開きます。このファイルには、アドインの作業ウインドウでレンダリングされる HTML が含まれています。

2.  `<body>` 要素を次のマークアップに置き換えます。

    ```html
    <body class="ms-font-m ms-welcome">
        <div id="content-header">
            <div class="padding">
                <h1>Welcome</h1>
            </div>
        </div>
        <div id="content-main">
            <div class="padding">
                <p>Select a task and then choose the buttons below and observe the output in the <b>Results</b> textbox.</p>
                <h3>Try it out</h3>
                <button class="ms-Button" id="get-task-guid">Get Task GUID</button>
                <br/><br/>
                <button class="ms-Button" id="get-task">Get Task data</button>
                <br/>
                <h4>Results:</h4>
                <textarea id="result" rows="6" cols="25"></textarea>
            </div>
        </div>
        <script type="text/javascript" src="node_modules/jquery/dist/jquery.js"></script>
        <script type="text/javascript" src="node_modules/office-ui-fabric-js/dist/js/fabric.js"></script>
    </body>
    ```

3.  **src/index.js** ファイルを開いて、アドインのスクリプトを指定します。 すべての内容を次のコードに置き換え、ファイルを保存します。

    ```js
    'use strict';

    (function () {

        var taskGuid;

        // The initialize function must be run each time a new page is loaded
        Office.initialize = function (reason) {
            $(document).ready(function () {
                $('#get-task-guid').click(getTaskGUID);
                $('#get-task').click(getTask);
            });
        };

        function getTaskGUID() {
            Office.context.document.getSelectedTaskAsync(function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    result.value = "Task GUID: " + asyncResult.value;
                    taskGuid = asyncResult.value;
                }
                else {
                    console.log(asyncResult.error.message);
                }
            });
        }

        function getTask() {
            if (taskGuid != undefined) {
                Office.context.document.getTaskAsync(
                    taskGuid,
                    function (asyncResult) {
                        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                            var taskInfo = asyncResult.value;
                            var taskOutput = "Task name: " + taskInfo.taskName +
                                            "\nGUID: " + taskGuid +
                                            "\nWSS Id: " + taskInfo.wssTaskId +
                                            "\nResource names: " + taskInfo.resourceNames;
                            result.value = taskOutput;
                        } else {
                            console.log(asyncResult.error.message);
                        }
                    }
                );
            } else {
                result.value = 'Task GUID not valid:\n' + taskGuid;
            } 
        }
    })();
    ```

4. プロジェクトのルートにある **app.css** ファイルを開いて、アドインのカスタム スタイルを指定します。 すべての内容を次の内容に置き換えて、ファイルを保存します。

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

## <a name="update-the-manifest"></a>マニフェストを更新する

1. ファイル **my-office-add-in-manifest.xml** ファイルを開いて、アドインの設定と機能を定義します。

2. `ProviderName` 要素にはプレースホルダーの値があります。これを自分の名前で置き換えます。

3.  `Description` 要素の `DefaultValue` には、プレースホルダーが含まれています。 それを **Project 用の作業ウィンドウ アドイン** に置き換えます。

4. ファイルを保存します。

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Project"/>
    ...
    ```

## <a name="start-the-dev-server"></a>開発用サーバーを起動する

[!include[Start server section](../includes/quickstart-yo-start-server.md)] 

## <a name="try-it-out"></a>お試しください。

1. Projectで、少なくとも 1 つのタスクを含む簡単なプロジェクトを作成します。

2. 使用するプラットフォームの手順に従って、アドインを実行し、プロジェクト内のアドインをサイドローディングします。

    - Windows: [Windows で Office アドインをサイドローディングします](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
    - Project Online: [Office Online で Office アドインをサイドローディングします](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)
    - iPad および Mac: [iPad と Mac で Office アドインをサイドローディングします](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)

3. Project のタスクを選択します。

    ![1 つのタスクが選択された Project でのプロジェクト計画のスクリーンショット](../images/project_quickstart_addin_1.png)

4. 作業ウィンドウで **[タスク GUID を取得]** を選択して、タスク GUID を **[結果]** テキストボックスに記入してください。

    ![1 つのタスクが選択された Project 内の Project 計画および作業ウィンドウのテキストボックスに記入されたタスク GUID のスクリーンショット](../images/project_quickstart_addin_2.png)

5. 作業ウィンドウで **[タスク データを取得]** 選択して、選択したタスクのいくつかのプロパティを **[結果]** テキストボックスに記入してください。

    ![1 つのタスクが選択された Project 内のプロジェクト計画および作業ウィンドウのテキストボックスに記入されたいくつかのタスクのプロパティのスクリーンショット](../images/project_quickstart_addin_3.png)

## <a name="next-steps"></a>次のステップ

おめでとうございます、Project アドインが正しく作成されました！ 次は、Project アドインの機能および一般的なシナリオについての説明です。

> [!div class="nextstepaction"]
> [Project アドイン](../project/project-add-ins.md)
