---
title: 最初の Project の作業ウィンドウ アドインを作成する
description: Office JS API を使用して単純な Project 作業ウィンドウ アドインを作成する方法について説明します。
ms.date: 08/04/2021
ms.prod: project
localization_priority: Priority
ms.openlocfilehash: 8cada7195d61b26867beeff87fc95ac2b2864c911b07a53540f1934fbaddc420
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2021
ms.locfileid: "57089430"
---
# <a name="build-your-first-project-task-pane-add-in"></a>最初の Project の作業ウィンドウ アドインを作成する

この記事では、Project の作業ウィンドウ アドインを作成するプロセスを紹介します。

## <a name="prerequisites"></a>前提条件

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]
[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

- Windows の Project 2016 またはそれ以降

## <a name="create-the-add-in"></a>アドインを作成する

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- **Choose a project type: (プロジェクトの種類を選択)** `Office Add-in Task Pane project`
- **Choose a script type: (スクリプトの種類を選択)** `Javascript`
- **What would you want to name your add-in?: (アドインの名前を何にしますか)** `My Office Add-in`
- **Which Office client application would you like to support?: (どの Office クライアント アプリケーションをサポートしますか)** `Project`

![コマンド ライン インターフェイスでの Yeoman ジェネレーターのプロンプトと回答を示すスクリーンショット。](../images/yo-office-project.png)

ウィザードを完了すると、ジェネレーターによってプロジェクトが作成されて、サポートしているノード コンポーネントがインストールされます。

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="explore-the-project"></a>プロジェクトを確認する

Yeomanジェネレーターで作成したアドインプロジェクトには、原型となる作業ペインアドインのサンプルコードが含まれています。

- プロジェクトのルート ディレクトリにある **./manifest.xml** ファイルで、アドインの機能と設定を定義します。
- **./src/taskpane/taskpane.html** ファイルには、作業ペイン用のHTMLマークアップが含まれています。
- **./src/taskpane/taskpane.css** ファイルには、作業ペインのコンテンツに適用されるCSSが含まれています。
- **./src/taskpane/taskpane.js** ファイルには、作業ウィンドウと Office クライアント アプリケーションの間のやり取りを容易にする Office JavaScript API コードが含まれています。

## <a name="update-the-code"></a>コードを更新する

コード エディターでファイル **./src/taskpane/taskpane.js** を開き、次のコードを `run` 関数内に追加します。 このコードでは、Office JavaScript API を使用して、選択したタスクの `Name`フィールドと `Notes` フィールドを設定します。

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

1. プロジェクトのルート フォルダーに移動します。

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

1. プロジェクトの依存関係をインストールします。

     ```command&nbsp;line
    npm install
    ```

1. ローカル Web サーバーを開始します。

    > [!NOTE]
    > 開発の最中でも、OfficeアドインはHTTPではなくHTTPSを使用する必要があります。 次のコマンドを実行した後に証明書をインストールするように求められた場合は、Yeoman ジェネレーターによって提供される証明書をインストールするプロンプトを受け入れます。

    プロジェクトのルート ディレクトリから次のコマンドを実行します。 このコマンドを実行すると、ローカル Web サーバーが起動します。

    ```command&nbsp;line
    npm run dev-server
    ```

1. Project で、簡素なプロジェクト計画を作成します。

1. [Windows に Office アドインをサイドロードする](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) の手順に従い、Project でアドインを読み込みます。

1. プロジェクト内の単一のタスクを選択します。

1. 作業ウィンドウの下部で **Run** リンクを選択して、 選択されたタスクの名前を変更し、そのタスクにメモを追加します。

    ![読み込まれた作業ウィンドウ アドインを用いた Project アプリケーションのスクリーンショット。](../images/project-quickstart-addin-1.png)

## <a name="next-steps"></a>次の手順

これで完了です。Project 作業ウィンドウのアドインが正常に作成されました。次に、Project アドインの機能を説明し、一般的なシナリオについて調べます。

> [!div class="nextstepaction"]
> [Project 用アドイン](../project/project-add-ins.md)

## <a name="see-also"></a>関連項目

- [Office アドインを開発する](../develop/develop-overview.md)
- [Office アドインの中心概念](../overview/core-concepts-office-add-ins.md)
