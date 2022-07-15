---
title: 最初の Project の作業ウィンドウ アドインを作成する
description: Office JS API を使用して単純な Project 作業ウィンドウ アドインを作成する方法について説明します。
ms.date: 07/13/2022
ms.prod: project
ms.localizationpriority: high
ms.openlocfilehash: c2f0e31b5a4c958cd155dfeb6d1648f7a2697c69
ms.sourcegitcommit: 9bb790f6264f7206396b32a677a9133ab4854d4e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/15/2022
ms.locfileid: "66797478"
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

![コマンド ライン インターフェイスに表示された Yeoman ジェネレーターのプロンプトと回答。](../images/yo-office-project.png)

ウィザードを完了すると、ジェネレーターによってプロジェクトが作成されて、サポートしているノード コンポーネントがインストールされます。

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="explore-the-project"></a>プロジェクトを確認する

Yeomanジェネレーターで作成したアドインプロジェクトには、原型となる作業ペインアドインのサンプルコードが含まれています。

- プロジェクトのルート ディレクトリにある **./manifest.xml** ファイルで、アドインの機能と設定を定義します。
- **./src/taskpane/taskpane.html** ファイルには、作業ペイン用のHTMLマークアップが含まれています。
- **./src/taskpane/taskpane.css** ファイルには、作業ペインのコンテンツに適用されるCSSが含まれています。
- **./src/taskpane/taskpane.js** ファイルには、作業ウィンドウと Office クライアント アプリケーションの間のやり取りを容易にする Office JavaScript API コードが含まれています。 このクイック スタートのコードは、選択したプロジェクト タスクの `Name` フィールドと `Notes` フィールドを設定します。

## <a name="try-it-out"></a>試してみる

1. プロジェクトのルート フォルダーに移動します。

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

1. ローカル Web サーバーを開始します。

    [!INCLUDE [alert use https](../includes/alert-use-https.md)]

    プロジェクトのルート ディレクトリから次のコマンドを実行します。 このコマンドを実行すると、ローカル Web サーバーが起動します。

    ```command&nbsp;line
    npm run dev-server
    ```

1. Project で、簡素なプロジェクト計画を作成します。

1. [Windows に Office アドインをサイドロードする](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) の手順に従い、Project でアドインを読み込みます。

1. プロジェクト内の単一のタスクを選択します。

1. 作業ウィンドウの下部で **Run** リンクを選択して、 選択されたタスクの名前を変更し、そのタスクにメモを追加します。

    ![Project アプリケーション上に読み込まれた作業ウィンドウ アドイン。](../images/project-quickstart-addin-1.png)

## <a name="next-steps"></a>次の手順

これで完了です。Project 作業ウィンドウのアドインが正常に作成されました。次に、Project アドインの機能を説明し、一般的なシナリオについて調べます。

> [!div class="nextstepaction"]
> [Project 用アドイン](../project/project-add-ins.md)

## <a name="see-also"></a>関連項目

- [Office アドインを開発する](../develop/develop-overview.md)
- [Office アドインの中心概念](../overview/core-concepts-office-add-ins.md)
- [Visual Studio コードを使用して発行する](../publish/publish-add-in-vs-code.md#using-visual-studio-code-to-publish)
