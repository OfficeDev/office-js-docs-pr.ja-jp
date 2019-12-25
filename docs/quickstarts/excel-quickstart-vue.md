---
title: Vue を使用して Excel 作業ウィンドウ アドインを作成する
description: ''
ms.date: 12/24/2019
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: a8ba3ba1c401e1433eb5be121ea37b053b1a4896
ms.sourcegitcommit: 350f5c6954dec3e9384e2030cd3265aaba7ae904
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/23/2019
ms.locfileid: "40851328"
---
# <a name="build-an-excel-task-pane-add-in-using-vue"></a>Vue を使用して Excel 作業ウィンドウ アドインを作成する

この記事では、Vue と Excel JavaScript API を使用して Excel 作業ウィンドウ アドインを構築するプロセスについて説明します。

## <a name="prerequisites"></a>前提条件

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

- [Vue CLI](https://cli.vuejs.org/) をグローバルにインストールします。

  ```command&nbsp;line
  npm install -g @vue/cli
  ```

## <a name="generate-a-new-vue-app"></a>新しい Vue アプリの生成

Vue CLI を使用して新しい Vue アプリを生成します。 端末から次のコマンドを実行します。

```command&nbsp;line
vue create my-add-in
```

次に、`default` プリセットを選択します。 Yarn または NPM のいずれかをパッケージとして使用するように求められたら、どちらかを選択できます。

## <a name="generate-the-manifest-file"></a>マニフェスト ファイルを生成する

各アドインには、設定と機能を定義するマニフェスト ファイルが必要です。

1. アプリ フォルダーに移動します。

    ```command&nbsp;line
    cd my-add-in
    ```

2. 以下のコマンドを実行し、Yeoman ジェネレーター使用してアドインのマニフェスト ファイルを生成します。

    ```command&nbsp;line
    yo office
    ```

    > [!NOTE]
    > `yo office`コマンドを実行すると、Yeoman のデータ収集ポリシーと Office アドイン CLI ツールに関するプロンプトが表示される場合があります。 提供された情報を使用して、必要に応じてプロンプトに応答します。 2 番目のプロンプトに対して [**終了**] を選択した場合、アドイン プロジェクトを作成する準備ができたら`yo office`コマンドを再度実行する必要があります。

    プロンプトが表示されたら、以下の情報を入力してアドイン プロジェクトを作成します。

    - **Choose a project type: (プロジェクトの種類を選択)** `Office Add-in project containing the manifest only`
    - **What would you want to name your add-in?: (アドインの名前を何にしますか)** `my-office-add-in`
    - **Which Office client application would you like to support?: (どの Office クライアント アプリケーションをサポートしますか)** `Excel`

    ![Yeoman ジェネレーター](../images/yo-office-manifest-only-vue.png)

ウィザードを完了すると、`my-office-add-in`フォルダーが`manifest.xml`ファイルを含んで作成されます。 マニフェストを使用して、クイック スタートの最後にアドインをサイドロードおよびテストします。

> [!TIP]
> アドイン プロジェクトの作成後に Yeoman ジェネレーターが提供する*次の手順*ガイダンスは無視できます。 この記事中の詳しい手順は、このチュートリアルを完了するために必要なすべてのガイダンスを提供します。

## <a name="secure-the-app"></a>アプリをセキュリティ保護する

[!include[HTTPS guidance](../includes/https-guidance.md)]

アプリで HTTPS を有効にするには、Vue プロジェクトのルート フォルダーに次の内容で `vue.config.js` ファイルを作成します。

```js
module.exports = {
  devServer: {
    port: 3000,
    https: true
  }
};
```

## <a name="update-the-app"></a>アプリを更新する

1. `public/index.html` ファイルを開き、`</head>` タグの直前に次の `<script>` タグを追加します。

   ```html
   <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
   ```

2. `src/main.js` を開き、内容を次のコードで置き換えます。

   ```js
   import Vue from 'vue';
   import App from './App.vue';

   Vue.config.productionTip = false;

   window.Office.initialize = () => {
     new Vue({
       render: h => h(App)
     }).$mount('#app');
   };
   ```

3. `src/App.vue` を開き、ファイル内容を次のコードで置き換えます。

   ```html
   <template>
     <div id="app">
       <div class="content">
         <div class="content-header">
           <div class="padding">
             <h1>Welcome</h1>
           </div>
         </div>
         <div id="content-main">
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

1. ターミナルから、次のコマンドを実行してデベロッパー サーバーを起動します。

   ```command&nbsp;line
   npm run serve
   ```

2. Web ブラウザーで `https://localhost:3000` (`https` に注意) に移動します。 ブラウザーにサイトの証明書が信頼されていないことが示された場合は、その[証明書を信頼するようコンピューターを構成する](https://github.com/OfficeDev/generator-office/blob/fd600bbe00747e64aa5efb9846295a3f66d428aa/src/docs/ssl.md#add-certification-file-through-ie)必要があります。

3. `https://localhost:3000` のページが空白で、証明書エラーがない場合、それは機能していることを意味します。 Vue アプリは、Office の初期化後にマウントされるため、Excel 環境内のもののみを表示します。

## <a name="try-it-out"></a>試してみる

1. アドインを実行して、Excel 内のアドインをサイドロードするのに使用するプラットフォームの手順に従います。

   - Windows: [Windows で Office アドインをサイドロードする](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
   - Web ブラウザー: [Office on the web で Office アドインをサイドロードする](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web)
   - iPad および Mac: [iPad と Mac で Office アドインをサイドロードする](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)

2. Excel で、**[ホーム]** タブを選択し、リボンの **[作業ウィンドウの表示]** ボタンをクリックして、アドインの作業ウィンドウを開きます。

   ![Excel アドイン ボタン](../images/excel-quickstart-addin-2a.png)

3. ワークシート内で任意のセルの範囲を選択します。

4. 作業ウィンドウで、**[色の設定]** ボタンをクリックして、選択範囲の色を緑に設定します。

   ![Excel アドイン](../images/excel-quickstart-addin-2c.png)

## <a name="next-steps"></a>次の手順

おめでとうございます! これで Vue を使用して Excel 作業ウィンドウ アドインを作成できました。 次に、Excel アドインの機能の詳細について説明します。Excel アドインのチュートリアルに従って、より複雑なアドインをビルドします。

> [!div class="nextstepaction"]
> [Excel アドインのチュートリアル](../tutorials/excel-tutorial.md)

## <a name="see-also"></a>関連項目

* [Office アドイン プラットフォームの概要](../overview/office-add-ins.md)
* [Office アドインを構築する](../overview/office-add-ins-fundamentals.md)
* [Office アドインを開発する](../develop/develop-overview.md)
* [Excel JavaScript API を使用した基本的なプログラミングの概念](../excel/excel-add-ins-core-concepts.md)
* [Excel アドインのコード サンプル](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [Excel JavaScript API リファレンス](/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
