# <a name="build-your-first-project-add-in"></a>最初の Project アドインを作成する

この記事では、jQuery と Office JavaScript API を使用して Project アドインを作成する手順について説明します。

## <a name="prerequisites"></a>前提条件

- [Node.js](https://nodejs.org)

- [Yeoman](https://github.com/yeoman/yo) と [Office アドイン用の Yeoman ジェネレーター](https://github.com/OfficeDev/generator-office) の最新バージョンをインストールします。


    ```bash
    npm install -g yo generator-office
    ```

## <a name="create-the-add-in"></a>アドインを作成する

1. ローカルドライブにフォルダーを作成し、`my-project-addin`という名前を付けてください。ここにアドインのファイルを作成します。

    ```bash
    mkdir my-project-addin
    ```

2. 新しいフォルダーに移動してください。

    ```bash
    cd my-project-addin
    ```

3. Yeomange ジェネレータを使用して Project アドインのプロジェクトを作成してください。次のコマンドを実行し、プロンプトに次のように応答します。

    ```bash
    yo office
    ```

    - **プロジェクトタイプを選択してください** `Office Add-in project using Jquery framework`
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

1. コード エディターで、プロジェクトのルートに**index.html**を開きます。このファイルには、アドインの作業ウインドウでレンダリングされる HTML が含まれています。

2. 要素`<body>`を次のマークアップに置き換えます。

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

3. **src/index.js**ファイルを開いて、アドインのスクリプトを特定します。 すべての内容を次のコードに置き換え、ファイルを保存します。

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

1. **manifest.xml** ファイルを開いて、アドインの設定と機能を定義します。

2. `ProviderName`要素にはプレースホルダー値が含まれています。 それを自分の名前に置き換えます。

3. 要素 `DefaultValue` の属性 `Description` にはプレースホルダーが含まれています。 それを ** Project 用の作業ウィンドウ アドイン** に置き換えます。

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

1. プロジェクトには、少なくとも 1 つのタスクを含む簡単なプロジェクトを作成してください。

2. アドインを実行するのに使用するプラットフォームの手順に従い、プロジェクト 内でアドインをサイドロードします。

    - Windows: [Windows で Office アドインをサイドロードする](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
    - Project Online:[Office Online で Office アドインをサイドロードする](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)
    - iPad および Mac: [iPad と Mac で Office アドインをサイドロードする](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)

3. プロジェクトには、タスクを選択してください。

    ![1 つのタスクが選択されたプロジェクトのプロジェクト計画のスクリーンショット](../images/project_quickstart_addin_1.png)

4. 作業ウィンドウで **[タスク GUID を取得]** ボタンを選択して、タスク GUID を **[結果]** テキストボックスに記入してください。

    ![1 つのタスクが選択された Project のプロジェクト計画および作業ウィンドウのテキストボックスに記入されたタスク GUID のスクリーンショット](../images/project_quickstart_addin_2.png)

5. 作業ウィンドウで **[タスク データを取得]** ボタンを選択して、選択したタスクのいくつかのプロパティを **[結果]** テキストボックスに記入してください。

    ![1 つのタスクが選択された Project のプロジェクト計画および作業ウィンドウのテキストボックスに記入された複数のタスクのプロパティのスクリーンショット](../images/project_quickstart_addin_3.png)

## <a name="next-steps"></a>次の手順

完了しました。 Project アドインが正常に作成されました！ 次に、 Project アドインの機能と一般的なシナリオについて調べます。

> [!div class="nextstepaction"]
> [  Project アドイン](../project/project-add-ins.md)
