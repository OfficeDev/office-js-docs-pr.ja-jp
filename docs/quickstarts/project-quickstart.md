---
title: 最初の Project の作業ウィンドウ アドインを作成する
description: ''
ms.date: 05/08/2019
ms.prod: project
localization_priority: Priority
ms.openlocfilehash: d61f8d83b88dbe69ff0ba9cd4b0afef77a4f03d6
ms.sourcegitcommit: a99be9c4771c45f3e07e781646e0e649aa47213f
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/11/2019
ms.locfileid: "33952252"
---
# <a name="build-your-first-project-task-pane-add-in"></a>最初の Project の作業ウィンドウ アドインを作成する

この記事では、Project の作業ウィンドウ アドインを作成するプロセスを紹介します。

## <a name="prerequisites"></a>前提条件

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

- Windows の Project 2016 またはそれ以降

## <a name="create-the-add-in"></a>アドインを作成する

1. Yeoman ジェネレーターを使用して、Project アドイン プロジェクトを作成します。 次のコマンドを実行し、以下のプロンプトに応答します。

    ```command&nbsp;line
    yo office
    ```

    - **Choose a project type: (プロジェクトの種類を選択)** `Office Add-in Task Pane project`
    - **Choose a script type: (スクリプトの種類を選択)** `Javascript`
    - **What would you want to name your add-in?: (アドインの名前を何にしますか)** `My Office Add-in`
    - **Which Office client application would you like to support?: (どの Office クライアント アプリケーションをサポートしますか)** `Project`

    ![Yeoman ジェネレーターのプロンプトと応答のスクリーンショット](../images/yo-office-project.png)
    
    ウィザードを完了すると、ジェネレーターによってプロジェクトが作成されて、サポートしているノード コンポーネントがインストールされます。
    
2. プロジェクトのルート フォルダーに移動します。

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

## <a name="explore-the-project"></a>プロジェクトを確認する

Yeoman ジェネレーターで作成したアドイン プロジェクトには、とても基本的な作業ウィンドウ アドインのサンプル コードが含まれています。 

- プロジェクトのルート ディレクトリにある **./manifest.xml**ファイルで、アドインの機能と設定を定義します。
- **./src/taskpane/taskpane.html**ファイルには、作業ウィンドウ用の HTML マークアップが含まれています。
- **./src/taskpane/taskpane.css**ファイルには、作業ウィンドウ内のコンテンツに適用される CSS が含まれています。
- **./src/taskpane/taskpane.js**ファイルには、作業ウィンドウと Office のホスト アプリケーションの間のやり取りを容易にする Office JavaScript API コードが含まれています。

## <a name="update-the-code"></a>コードを更新する

コード エディターでファイル **./src/taskpane/taskpane.js** を開き、次のコードを **実行** 関数内に追加します。 このコードでは、Office JavaScript API を使用して、選択したタスクの `Name`フィールドと `Notes` フィールドを設定します。

```js
var taskGuid;

// Get the GUID of the selected task
Office.context.document.getSelectedTaskAsync(
    function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            taskGuid = result.value;

            // Set the specified fields for the selected task.
            var targetFields = [Office.ProjectTaskFields.Name, Office.ProjectTaskFields.Notes];
            var fieldValues = ['New task name', 'Notes for the task.'];

            // Set the field value. If the call is successful, set the next field.
            for (var i = 0; i < targetFields.length; i++) {
                Office.context.document.setTaskFieldAsync(
                    taskGuid,
                    targetFields[i],
                    fieldValues[i],
                    function (result) {
                        if (result.status === Office.AsyncResultStatus.Succeeded) {
                            i++;
                        }
                        else {
                            var err = result.error;
                            console.log(err.name + ' ' + err.code + ' ' + err.message);
                        }
                    }
                );
            }
        } else {
            var err = result.error;
            console.log(err.name + ' ' + err.code + ' ' + err.message);
        }
    }
);
```

## <a name="try-it-out"></a>試してみる

1. 次のコマンドを実行してローカル Web サーバーを起動します:

    ```command&nbsp;line
    npm start
    ```

    > [!NOTE]
    > Office アドインでは、開発中であっても HTTP ではなく HTTPS を使用する必要があります。 `npm start`を実行した後に証明書をインストールするように求められた場合は、Yeoman ジェネレーターによって提供される証明書をインストールするプロンプトを受け入れます。 

2. Project で、簡素なプロジェクト計画を作成します。

3. [Windows に Office アドインをサイドロードする](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) の手順に従い、Project でアドインを読み込みます。

4. プロジェクト内の単一のタスクを選択します。

5. 作業ウィンドウの下部で **Run** リンクを選択して、 選択されたタスクの名前を変更し、そのタスクにメモを追加します。

    ![読み込まれた作業ウィンドウ アドインを用いた Project アプリケーションのスクリーンショット](../images/project-quickstart-addin-1.png)

## <a name="next-steps"></a>次の手順

おめでとうございます。 Project の作業ウィンドウ アドインが正常に作成されました。 この後は、Project アドインの機能と一般的なシナリオについて調べます。

> [!div class="nextstepaction"]
> [Project 用アドイン](../project/project-add-ins.md)

