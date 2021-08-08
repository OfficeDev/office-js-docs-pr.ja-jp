---
title: Vue を使用して Excel 作業ウィンドウ アドインを構築する
description: Office JS API と Vue を使用して単純な Excel 作業ウィンドウ アドインを作成する方法について説明します。
ms.date: 08/04/2021
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 1686f9d9537718eb5ba56fa9ea7f0b4ccb7d65ec
ms.sourcegitcommit: e570fa8925204c6ca7c8aea59fbf07f73ef1a803
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/05/2021
ms.locfileid: "53774442"
---
# <a name="use-vue-to-build-an-excel-task-pane-add-in"></a>Vue を使用して Excel 作業ウィンドウ アドインを構築する

この記事では、Vue と Excel JavaScript API を使用して Excel 作業ウィンドウ アドインを構築するプロセスについて説明します。

## <a name="prerequisites"></a>前提条件

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]
[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

- [Vue CLI](https://cli.vuejs.org/) をグローバルにインストールします。 ターミナルから、次のコマンドを実行します。

  ```command&nbsp;line
  npm install -g @vue/cli
  ```

## <a name="generate-a-new-vue-app"></a>新しい Vue アプリの生成

新しい Vue アプリを生成するには、Vue CLI を使用します。

```command&nbsp;line
vue create my-add-in
```

次に、「Vue 3」の `Default` プリセットを選択します (お好みで「Vue 2」を選択します)。

## <a name="generate-the-manifest-file"></a>マニフェスト ファイルを生成する

各アドインには、設定と機能を定義するマニフェスト ファイルが必要です。

1. アプリ フォルダーに移動します。

    ```command&nbsp;line
    cd my-add-in
    ```

1. Yeoman ジェネレーター使用して、アドインのマニフェスト ファイルを生成します。

    ```command&nbsp;line
    yo office
    ```

    > [!NOTE]
    > `yo office`コマンドを実行すると、Yeoman のデータ収集ポリシーと Office アドイン CLI ツールに関するプロンプトが表示される場合があります。 提供された情報を使用して、必要に応じてプロンプトに応答します。 2 番目のプロンプトに対して [**終了**] を選択した場合、アドイン プロジェクトを作成する準備ができたら`yo office`コマンドを再度実行する必要があります。

    プロンプトが表示されたら、以下の情報を入力してアドイン プロジェクトを作成します。

    - **Choose a project type: (プロジェクトの種類を選択)** `Office Add-in project containing the manifest only`
    - **What would you want to name your add-in?: (アドインの名前を何にしますか)** `My Office Add-in`
    - **Which Office client application would you like to support?: (どの Office クライアント アプリケーションをサポートしますか)** `Excel`

    ![プロジェクトの種類がマニフェスト専用に設定されている Yeoman Office アドイン ジェネレーター コマンドライン インターフェイスのスクリーンショット。](../images/yo-office-manifest-only-vue.png)

完了後、ウィザードは **manifest.xml** ファイルを含む **個人用 Office アドイン** フォルダーを作成します。 マニフェストを使用して、アドインをサイドロードしてテストします。

> [!TIP]
> アドイン プロジェクトの作成後に Yeoman ジェネレーターが提供する *次の手順* ガイダンスは無視できます。 この記事中の詳しい手順は、このチュートリアルを完了するために必要なすべてのガイダンスを提供します。

## <a name="secure-the-app"></a>アプリをセキュリティ保護する

[!include[HTTPS guidance](../includes/https-guidance.md)]

1. アプリの HTTPS を有効にします。 Vue プロジェクトのルート フォルダーに次の内容で **vue.config.js** ファイルを作成します。

    ```js
    var fs = require("fs");
    var path = require("path");
    var homedir = require('os').homedir()
  
    module.exports = {
      devServer: {
        port: 3000,
        https: true,
        key: fs.readFileSync(path.resolve(`${homedir}/.office-addin-dev-certs/localhost.key`)),
        cert: fs.readFileSync(path.resolve(`${homedir}/.office-addin-dev-certs/localhost.crt`)),
        ca: fs.readFileSync(path.resolve(`${homedir}/.office-addin-dev-certs/ca.crt`))
      }
    }
    ```

1. アドインの証明書をインストールします。

   ```command&nbsp;line
   npx office-addin-dev-certs install
   ```

## <a name="explore-the-project"></a>プロジェクトを探究する

Yeoman ジェネレーターで作成したアドイン プロジェクトには、基本的なアドインの作業ウィンドウが含まれています。 アドイン プロジェクトの主要な構成要素を確認したい場合は、コード エディターでプロジェクトを開き、以下に一覧表示されているファイルを確認します。 アドインを試す準備ができたら、次のセクションに進みます。

- プロジェクトのルート ディレクトリにある **manifest.xml** ファイルで、アドインの機能と設定を定義します。 **manifest.xml** ファイルの詳細については、「[Office アドインの XML マニフェスト](../develop/add-in-manifests.md)」を参照してください。
- **./src/App.vue** ファイルには、作業ウィンドウの HTML マークアップ、作業ウィンドウのコンテンツに適用される CSS、作業ウィンドウと Excel の間の対話操作を容易にする Office JavaScript API コードが含まれます。

## <a name="update-the-app"></a>アプリを更新する

1. **./public/index.html** ファイルを開き、`</head>` タグの直前に次の `<script>` タグを追加します。

   ```html
   <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
   ```

1. **manifest.xml** を開き、`<Resources>` タグの中で `<bt:Urls>` タグを検索します。 ID `Taskpane.Url` を持つ `<bt:Url>` タグを検索し、その `DefaultValue` 属性を更新します。 新しい `DefaultValue` は `https://localhost:3000/index.html` です。 更新されたタグ全体が次の行と一致する必要があります。

   ```html
   <bt:Url id="Taskpane.Url" DefaultValue="https://localhost:3000/index.html" />
   ```

1. **./src/main.js** を開き、内容を次のコードで置き換えます。

   ```js
   import { createApp } from 'vue'
   import App from './App.vue'

   window.Office.onReady(() => {
       createApp(App).mount('#app');
   });
   ```

1. **./src/App.vue** を開き、ファイル内容を次のコードで置き換えます。

   ```html
   <template>
     <div id="app">
       <div class="content">
         <div class="content-header">
           <div class="padding">
             <h1>Welcome</h1>
           </div>
         </div>
         <div class="content-main">
           <div class="padding">
             <p>
               Choose the button below to set the color of the selected range to
               green.
             </p>
             <br />
             <h3>Try it out</h3>
             <button @click="onSetColor">Set color</button>
           </div>
         </div>
       </div>
     </div>
   </template>

   <script>
     export default {
       name: 'App',
       methods: {
         onSetColor() {
           window.Excel.run(async context => {
             const range = context.workbook.getSelectedRange();
             range.format.fill.color = 'green';
             await context.sync();
           });
         }
       }
     };
   </script>

   <style>
     .content-header {
       background: #2a8dd4;
       color: #fff;
       position: absolute;
       top: 0;
       left: 0;
       width: 100%;
       height: 80px;
       overflow: hidden;
     }

     .content-main {
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
   </style>
   ```

## <a name="start-the-dev-server"></a>開発用サーバーの起動

1. 依存関係をインストールします。

     ```command&nbsp;line
    npm install
    ```

1. 開発用サーバーを起動します。

   ```command&nbsp;line
   npm run serve
   ```

1. Web ブラウザーで `https://localhost:3000` (`https` に注意) に移動します。 `https://localhost:3000` のページが空白で、証明書エラーがない場合、それは機能していることを意味します。 Vue アプリは、Office の初期化後にマウントされるため、Excel 環境内のもののみを表示します。

## <a name="try-it-out"></a>試してみる

1. アドインを実行して、Excel 内のアドインをサイドロードします。 使用するプラットフォームの手順に従います。

   - Windows: [Windows で Office アドインをサイドロードする](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
   - Web ブラウザー: [Office on the web で Office アドインをサイドロードする](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web)
   - iPad および Mac: [iPad と Mac で Office アドインをサイドロードする](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)

1. Excel でアドイン作業ウィンドウを開きます。 **[ホーム]** タブで、**[作業ウィンドウの表示]** ボタンをクリックします。

   ![[作業ウィンドウの表示] ボタンが強調表示されている Excel ホーム メニューのスクリーンショット。](../images/excel-quickstart-addin-2a.png)

1. ワークシート内で任意のセルの範囲を選択します。

1. 選択範囲の色を緑に設定します。 アドインの作業ウィンドウで **[色の設定]** ボタンを選択します。

   ![アドイン作業ウィンドウを開いた状態の Excel のスクリーンショット。](../images/excel-quickstart-addin-2c.png)

## <a name="next-steps"></a>次の手順

これで完了です。Vue を使用して Excel タスク ウィンドウ アドインが正常に作成されました。次に、Excel アドインの機能の詳細について説明します。Excel アドインのチュートリアルに従って、より複雑なアドインをビルドします。

> [!div class="nextstepaction"]
> [Excel アドインのチュートリアル](../tutorials/excel-tutorial.md)

## <a name="see-also"></a>関連項目

- [Office アドイン プラットフォームの概要](../overview/office-add-ins.md)
- [Office アドインを開発する](../develop/develop-overview.md)
- [Office アドインの Excel JavaScript オブジェクト モデル](../excel/excel-add-ins-core-concepts.md)
- [Excel アドインのコード サンプル](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
- [Excel JavaScript API リファレンス](../reference/overview/excel-add-ins-reference-overview.md)
