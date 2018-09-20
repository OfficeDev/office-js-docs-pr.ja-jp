# <a name="build-your-first-project-add-in"></a>最初の Project アドインをビルドする

この記事では、jQuery と Office JavaScript API を使用して Project アドインを作成する手順について説明します。

## <a name="prerequisites"></a>前提条件

- [Node.js](https://nodejs.org)

- [Yeoman](https://github.com/yeoman/yo) の最新バージョンと [Office アドイン用の Yeoman ジェネレーター](https://github.com/OfficeDev/generator-office)をグローバルにインストールします。

    ```bash
    npm install -g yo generator-office
    ```

## <a name="create-the-add-in"></a>アドインを作成する

1. ローカル ドライブにフォルダーを作成し、`my-project-addin` という名前を付けます。 ここにアドインのファイルを作成します。

2. 新しいフォルダーに移動します。

    ```bash
    cd my-project-addin
    ```

3. Yeoman ジェネレーターを使用して、Project アドイン プロジェクトを作成します。 次のコマンドを実行し、以下のプロンプトに応答します。

    ```bash
    yo office
    ```

    - **Choose a project type:​ (プロジェクト タイプを選択してください)** `Office Add-in project using Jquery framework`
    - **Choose a script type: (スクリプト タイプを選択してください)** `Javascript`
    - **What would you want to name your add-in?: (アドインの名前を何にしますか)** `My Office Add-in`
    - **Which Office client application would you like to support?: (どの Office クライアント アプリケーションをサポートしますか)** `Project`

    ![Yeoman ジェネレーターのプロンプトと応答のスクリーンショット](../images/yo-office-project-jquery.png)
    
    ウィザードが完了すると、ジェネレーターはプロジェクトを作成し、サポートする Node コンポーネントをインストールします。
    
4. Web アプリケーション プロジェクトのルート フォルダーに移動します。

    ```bash
    cd "My Office Add-in"
    ```

## <a name="update-the-code"></a>コードを更新する

1. コード エディターで、プロジェクトのルートにある **index.html** を開きます。 このファイルには、アドインの作業ウィンドウにレンダリングされる HTML が含まれています。

2. 要素内の `<header>` 要素を次のマークアップに置き換えます。`<body>`

    ```html
    <div id="content-header">
        <div class="padding">
            <h1>Welcome</h1>
        </div>
    </div>
    ```

3. `<body>` 要素内の `<main>` 要素を次のマークアップに置き換えて、ファイルを保存します。

    ```html
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
    ```

4. **src/index.js** ファイルを開いて、アドインのスクリプトを指定します。 すべての内容を次のコードに置き換え、ファイルを保存します。

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

1. **my-office-add-in-manifest.xml** ファイルを開いて、アドインの設定と機能を定義します。

2. 要素にはプレースホルダー値が含まれています。`ProviderName` それを自分の名前に置き換えます。

3. 要素の `DefaultValue` 属性にはプレースホルダー値が含まれています。`Description` これは、**A task pane add-in for Project** に置き換えてください。

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

## <a name="start-the-dev-server"></a>開発用サーバーの起動

[!include[Start server section](../includes/quickstart-yo-start-server.md)] 

## <a name="try-it-out"></a>試してみる

1. 少なくとも 1 つのタスクを含むシンプルなプロジェクトを Project で作成します。

2. アドインを実行して、Project 内のアドインをサイドロードするのに使用するプラットフォームの手順に従います。

    - Windows: [Windows で Office アドインをサイドロードする](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
    - Project Online:[Office Online で Office アドインをサイドロードする](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)
    - iPad および Mac: [iPad と Mac で Office アドインをサイドロードする](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)

3. Project でタスクを選択します。

    ![1 つのタスクが選択された Project のプロジェクト計画のスクリーンショット](../images/project_quickstart_addin_1.png)

4. 作業ウィンドウで **[タスク GUID を取得]** ボタンを選択して、タスク GUID を **[結果]** テキストボックスに記入します。

    ![1 つのタスクが選択された Project のプロジェクト計画および作業ウィンドウのテキストボックスに記入されたタスク GUID のスクリーンショット](../images/project_quickstart_addin_2.png)

5. 作業ウィンドウで **[タスク データを取得]** ボタンを選択して、選択したタスクのいくつかのプロパティを **[結果]** テキストボックスに記入します。

    ![1 つのタスクが選択された Project のプロジェクト計画および作業ウィンドウのテキストボックスに記入された複数のタスクのプロパティのスクリーンショット](../images/project_quickstart_addin_3.png)

## <a name="next-steps"></a>次の手順

これで完了です。Project アドインが正常に作成されました。 この後は、Project アドインの機能と一般的なシナリオについて調べます。

> [!div class="nextstepaction"]
> [Project 用アドイン](../project/project-add-ins.md)
