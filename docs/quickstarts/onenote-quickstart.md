---
title: 最初の OneNote の作業ウィンドウ アドインを作成する
description: Office JS API を使用して単純な OneNote 作業ウィンドウ アドインを作成する方法について説明します。
ms.date: 08/04/2021
ms.prod: onenote
localization_priority: Priority
ms.openlocfilehash: a431b48eed7cedee75b7ef4121cec43d1777f5ecad06536752c1abedc92af725
ms.sourcegitcommit: f5d4321763e366a10f2d868fb329dbef5239c830
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/11/2021
ms.locfileid: "57845656"
---
# <a name="build-your-first-onenote-task-pane-add-in"></a>最初の OneNote の作業ウィンドウ アドインを作成する

この記事では、OneNote の作業ウィンドウ アドインを作成するプロセスを紹介します。

## <a name="prerequisites"></a>必須条件

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]
[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## <a name="create-the-add-in-project"></a>アドイン プロジェクトの作成

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- **Choose a project type: (プロジェクトの種類を選択)** `Office Add-in Task Pane project`
- **Choose a script type: (スクリプトの種類を選択)** `Javascript`
- **What would you want to name your add-in?: (アドインの名前を何にしますか)** `My Office Add-in`
- **Which Office client application would you like to support?: (どの Office クライアント アプリケーションをサポートしますか)** `OneNote`

![コマンド ライン インターフェイスでの Yeoman ジェネレーターのプロンプトと回答を示すスクリーンショット。](../images/yo-office-onenote.png)

ウィザードを完了すると、ジェネレーターによってプロジェクトが作成されて、サポートしているノード コンポーネントがインストールされます。

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="explore-the-project"></a>プロジェクトを確認する

Yeomanジェネレーターで作成したアドインプロジェクトには、原型となる作業ペインアドインのサンプルコードが含まれています。

- プロジェクトのルート ディレクトリにある **./manifest.xml** ファイルで、アドインの機能と設定を定義します。
- **./src/taskpane/taskpane.html** ファイルには、作業ペイン用のHTMLマークアップが含まれています。
- **./src/taskpane/taskpane.css** ファイルには、作業ペインのコンテンツに適用されるCSSが含まれています。
- **./src/taskpane/taskpane.js** ファイルには、作業ウィンドウと Office クライアント アプリケーションの間のやり取りを容易にする Office JavaScript API コードが含まれています。

## <a name="update-the-code"></a>コードを更新する

コード エディターでファイル **./src/taskpane/taskpane.js** を開き、次のコードを `run` 関数内に追加します。 このコードは、OneNote JavaScript API を使用してページ タイトルを設定し、ページの本文にアウトラインを追加します。

```js
try {
    await OneNote.run(async context => {

        // Get the current page.
        var page = context.application.getActivePage();

        // Queue a command to set the page title.
        page.title = "Hello World";

        // Queue a command to add an outline to the page.
        var html = "<p><ol><li>Item #1</li><li>Item #2</li></ol></p>";
        page.addOutline(40, 90, html);

        // Run the queued commands, and return a promise to indicate task completion.
        return context.sync();
    });
} catch (error) {
    console.log("Error: " + error);
}
```

## <a name="try-it-out"></a>試してみる

1. プロジェクトのルート フォルダーに移動します。

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

1. ローカル Web サーバーを起動し、アドインのサイドロードを行います。

    > [!NOTE]
    > Office アドインは、開発中であっても HTTP ではなく HTTPS を使用する必要があります。 次のいずれかのコマンドを実行した後に証明書をインストールするように求められた場合は、Yeoman ジェネレーターによって提供される証明書をインストールするプロンプトを受け入れます。 変更を行うには、管理者としてコマンド プロンプトまたはターミナルを実行する必要がある場合もあります。

    > [!TIP]
    > Mac でアドインをテストしている場合は、先に進む前に次のコマンドを実行してください。 このコマンドを実行すると、ローカル Web サーバーが起動します。
    >
    > ```command&nbsp;line
    > npm run dev-server
    > ```

    プロジェクトのルート ディレクトリから次のコマンドを実行します。 このコマンドを実行すると、ローカル Web サーバーが起動します (まだ実行されていない場合)。

    ```command&nbsp;line
    npm run start:web
    ```

1. [OneNote on the web](https://www.onenote.com/notebooks) でノートブックを開き、新しいページを作成します。

1. **[挿入] > [Office アドイン]** の順に選択し、[Office アドイン] ダイアログを開きます。

    - コンシューマー アカウントでサインインしている場合は、**[マイ アドイン]** タブを選択し、**[マイ アドインのアップロード]** を選択します。

    - 職場または教育用アカウントでサインインしている場合は、**[自分の所属組織]** タブを選択し、**[マイ アドインのアップロード]** を選択します。

    次の図は、コンシューマー ノートブックの **[マイ アドイン]** タブを示しています。

    ![[マイ アドイン] タブを示す [Office アドイン] ダイアログのスクリーンショット。](../images/onenote-office-add-ins-dialog.png)

1. [アドインのアップロード] ダイアログで、プロジェクト フォルダー内の **manifest.xml** を参照し、**[アップロード]** を選択します。

1. **[ホーム]** タブから、リボンの **[作業ウィンドウの表示]** ボタンをクリックします。 アドインの作業ウィンドウは、OneNote ページの横にある iFrame で開きます。

1. 作業ウィンドウの下部にある [**実行**] リンクをクリックしてページ タイトルを設定し、ページの本文にアウトラインを追加します。

    ![このウォークスルーから作成されたアドインを示すスクリーンショット: OneNote で作業ウィンドウのリボン ボタンと作業ウィンドウを表示します。](../images/onenote-first-add-in-4.png)

## <a name="next-steps"></a>次の手順

おめでとうございます。OneNote の作業ウィンドウ アドインが正常に作成されました。次に、OneNote アドイン構築の中心概念の詳細について説明します。

> [!div class="nextstepaction"]
> [OneNote の JavaScript API のプログラミングの概要](../onenote/onenote-add-ins-programming-overview.md)

## <a name="see-also"></a>関連項目

- [Office アドイン プラットフォームの概要](../overview/office-add-ins.md)
- [Office アドインを開発する](../develop/develop-overview.md)
- [OneNote の JavaScript API のプログラミングの概要](../onenote/onenote-add-ins-programming-overview.md)
- [OneNote JavaScript API リファレンス](../reference/overview/onenote-add-ins-javascript-reference.md)
- [Rubric Grader のサンプル](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
