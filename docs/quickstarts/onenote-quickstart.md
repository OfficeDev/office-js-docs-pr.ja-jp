---
title: 最初の OneNote の作業ウィンドウ アドインを作成する
description: ''
ms.date: 06/20/2019
ms.prod: onenote
localization_priority: Priority
ms.openlocfilehash: 24c8854cb1f9332371f3726409f91f7cdbf53243
ms.sourcegitcommit: 1dc1bb0befe06d19b587961da892434bd0512fb5
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/13/2019
ms.locfileid: "36308023"
---
# <a name="build-your-first-onenote-task-pane-add-in"></a>最初の OneNote の作業ウィンドウ アドインを作成する

この記事では、OneNote の作業ウィンドウ アドインを作成するプロセスを紹介します。

## <a name="prerequisites"></a>必須条件

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## <a name="create-the-add-in-project"></a>アドイン プロジェクトの作成

[!include[note about Yeoman generator bug](../includes/note-yeoman-generator-bug-201908.md)]

Yeoman ジェネレーターを使用して、OneNote アドイン プロジェクトを作成します。 次のコマンドを実行し、以下のプロンプトに応答します。

```command&nbsp;line
yo office
```

- **Choose a project type: (プロジェクトの種類を選択)** `Office Add-in Task Pane project`
- **Choose a script type: (スクリプトの種類を選択)** `Javascript`
- **What would you want to name your add-in?: (アドインの名前を何にしますか)** `my-office-add-in`
- **Which Office client application would you like to support?: (どの Office クライアント アプリケーションをサポートしますか)** `OneNote`

ウィザードを完了すると、ジェネレーターによってプロジェクトが作成されて、サポートしているノード コンポーネントがインストールされます。
    
## <a name="explore-the-project"></a>プロジェクトを確認する

Yeomanジェネレーターで作成したアドインプロジェクトには、原型となる作業ペインアドインのサンプルコードが含まれています。 

- プロジェクトのルートディレクトリにある **./ manifest.xml**ファイルは、アドインの設定と機能性を定義します。
- **./src/taskpane/taskpane.html**ファイルには、作業ペイン用のHTMLマークアップが含まれています。
- **./src/taskpane/taskpane.css**ファイルには、作業ウィンドウ内のコンテンツに適用される CSS が含まれています。
- **./src/taskpane/taskpane.js**ファイルには、作業ウィンドウと Office のホスト アプリケーションの間のやり取りを容易にする Office JavaScript API コードが含まれています。

## <a name="update-the-code"></a>コードを更新する

コード エディターでファイル **./src/taskpane/taskpane.js** を開き、次のコードを **実行** 関数内に追加します。 このコードは、OneNote JavaScript API を使用してページ タイトルを設定し、ページの本文にアウトラインを追加します。

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
    cd "my-office-add-in"
    ```

2. ローカル Web サーバーを起動し、アドインのサイドロードを行います。

    > [!NOTE]
    > Office アドインは、開発中であっても HTTP ではなく HTTPS を使用する必要があります。 次のいずれかのコマンドを実行した後に証明書をインストールするように求められた場合は、Yeoman ジェネレーターによって提供される証明書をインストールするプロンプトを受け入れます。

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

3. [OneNote on the web](https://www.onenote.com/notebooks) でノートブックを開き、新しいページを作成します。

4. **[挿入] > [Office アドイン]** の順に選択し、[Office アドイン] ダイアログを開きます。

    - コンシューマー アカウントでサインインしている場合は、**[マイ アドイン]** タブを選択し、**[マイ アドインのアップロード]** を選択します。

    - 職場または学校アカウントでサインインしている場合は、**[自分の所属組織]** タブを選択し、**[マイ アドインのアップロード]** を選択します。 

    次の図は、コンシューマー ノートブックの **[マイ アドイン]** タブを示しています。

    <img alt="The Office Add-ins dialog showing the MY ADD-INS tab" src="../images/onenote-office-add-ins-dialog.png" width="500">

5. [アドインのアップロード] ダイアログで、プロジェクト フォルダー内の **manifest.xml** を参照し、**[アップロード]** を選択します。 

6. **[ホーム]** タブから、リボンの **[作業ウィンドウの表示]** ボタンをクリックします。 アドインの作業ウィンドウは、OneNote ページの横にある iFrame で開きます。

7. 作業ウィンドウの下部にある [**実行**] リンクをクリックしてページ タイトルを設定し、ページの本文にアウトラインを追加します。

    ![このチュートリアルでビルドした OneNote アドイン](../images/onenote-first-add-in-4.png)

## <a name="next-steps"></a>次の手順

おめでとうございます。OneNote の作業ウィンドウ アドインが正常に作成されました。 次に、OneNote アドイン構築の中心概念の詳細について説明します。

> [!div class="nextstepaction"]
> [OneNote の JavaScript API のプログラミングの概要](../onenote/onenote-add-ins-programming-overview.md)

## <a name="see-also"></a>関連項目

- [OneNote の JavaScript API のプログラミングの概要](../onenote/onenote-add-ins-programming-overview.md)
- [OneNote JavaScript API リファレンス](/office/dev/add-ins/reference/overview/onenote-add-ins-javascript-reference)
- [Rubric Grader のサンプル](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Office アドイン プラットフォームの概要](../overview/office-add-ins.md)

