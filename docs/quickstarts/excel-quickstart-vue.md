---
title: Vue を使用して Excel 作業ウィンドウ アドインを作成する
description: ''
ms.date: 09/04/2019
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 9947852a586570345ba9f3dfe09340af6d01ace6
ms.sourcegitcommit: 78998a9f0ebb81c4dd2b77574148b16fe6725cfc
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/03/2019
ms.locfileid: "36715632"
---
# <a name="build-an-excel-task-pane-add-in-using-vue"></a><span data-ttu-id="1a51b-102">Vue を使用して Excel 作業ウィンドウ アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="1a51b-102">Build an Excel task pane add-in using Vue</span></span>

<span data-ttu-id="1a51b-103">この記事では、Vue と Excel JavaScript API を使用して Excel 作業ウィンドウ アドインを構築するプロセスについて説明します。</span><span class="sxs-lookup"><span data-stu-id="1a51b-103">In this article, you'll walk through the process of building an Excel task pane add-in using Vue and the Excel JavaScript API.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="1a51b-104">前提条件</span><span class="sxs-lookup"><span data-stu-id="1a51b-104">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

- <span data-ttu-id="1a51b-105">[Vue CLI](https://cli.vuejs.org/) をグローバルにインストールします。</span><span class="sxs-lookup"><span data-stu-id="1a51b-105">Install the [Vue CLI](https://cli.vuejs.org/) globally.</span></span>

  ```command&nbsp;line
  npm install -g @vue/cli
  ```

## <a name="generate-a-new-vue-app"></a><span data-ttu-id="1a51b-106">新しい Vue アプリの生成</span><span class="sxs-lookup"><span data-stu-id="1a51b-106">Generate a new Vue app</span></span>

<span data-ttu-id="1a51b-p101">Vue CLI を使用して新しい Vue アプリを生成します。 端末から次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="1a51b-p101">Use the Vue CLI to generate a new Vue app. From the terminal, run the following command and then answer the prompts as described below.</span></span>

```command&nbsp;line
vue create my-add-in
```

<span data-ttu-id="1a51b-109">次に、`default` プリセットを選択します。</span><span class="sxs-lookup"><span data-stu-id="1a51b-109">Then select the `default` preset.</span></span> <span data-ttu-id="1a51b-110">Yarn または NPM のいずれかをパッケージとして使用するように求められたら、どちらかを選択できます。</span><span class="sxs-lookup"><span data-stu-id="1a51b-110">If you are prompted to use either Yarn or NPM as a package you can choose either one.</span></span>

## <a name="generate-the-manifest-file"></a><span data-ttu-id="1a51b-111">マニフェスト ファイルを生成する</span><span class="sxs-lookup"><span data-stu-id="1a51b-111">Generate the manifest file</span></span>

<span data-ttu-id="1a51b-112">各アドインには、設定と機能を定義するマニフェスト ファイルが必要です。</span><span class="sxs-lookup"><span data-stu-id="1a51b-112">Each add-in requires a manifest file to define its settings and capabilities.</span></span>

1. <span data-ttu-id="1a51b-113">アプリ フォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="1a51b-113">Navigate to your app folder.</span></span>

   ```command&nbsp;line
   cd my-add-in
   ```

2. <span data-ttu-id="1a51b-p103">Yeoman ジェネレーター使用して、アドインのマニフェスト ファイルを生成します。 次のコマンドを実行し、以下に示すプロンプトに応答します。</span><span class="sxs-lookup"><span data-stu-id="1a51b-p103">Use the Yeoman generator to generate the manifest file for your add-in. Run the following command and then answer the prompts as shown below.</span></span>

   ```command&nbsp;line
   yo office
   ```

   ![Yeoman ジェネレーター](../images/yo-office-manifest-only-vue.png)

   - <span data-ttu-id="1a51b-117">**Choose a project type: (プロジェクトの種類を選択)** `Office Add-in project containing the manifest only`</span><span class="sxs-lookup"><span data-stu-id="1a51b-117">**Choose a project type:** `Office Add-in project containing the manifest only`</span></span>
   - <span data-ttu-id="1a51b-118">**What would you want to name your add-in?: (アドインの名前を何にしますか)**</span><span class="sxs-lookup"><span data-stu-id="1a51b-118">**What do you want to name your add-in?**</span></span> `my-office-add-in`
   - <span data-ttu-id="1a51b-119">**Which Office client application would you like to support?: (どの Office クライアント アプリケーションをサポートしますか)**</span><span class="sxs-lookup"><span data-stu-id="1a51b-119">**Which Office client application would you like to support?**</span></span> `Excel`

<span data-ttu-id="1a51b-120">ウィザードを完了すると、`manifest.xml` ファイルを含む `my-office-add-in` フォルダーが作成されます。</span><span class="sxs-lookup"><span data-stu-id="1a51b-120">After you complete the wizard, it creates a `my-office-add-in` folder, which contains a `manifest.xml` file.</span></span> <span data-ttu-id="1a51b-121">マニフェストを使用して、クイック スタートの最後にアドインをサイドロードおよびテストします。</span><span class="sxs-lookup"><span data-stu-id="1a51b-121">You will use the manifest to sideload and test your add-in at the end of the quick start.</span></span>

## <a name="secure-the-app"></a><span data-ttu-id="1a51b-122">アプリをセキュリティ保護する</span><span class="sxs-lookup"><span data-stu-id="1a51b-122">Secure the app</span></span>

[!include[HTTPS guidance](../includes/https-guidance.md)]

<span data-ttu-id="1a51b-123">アプリで HTTPS を有効にするには、Vue プロジェクトのルート フォルダーに次の内容で `vue.config.js` ファイルを作成します。</span><span class="sxs-lookup"><span data-stu-id="1a51b-123">To enable HTTPS for your app, create a `vue.config.js` file in the root folder of the Vue project with the following contents:</span></span>

```js
module.exports = {
  devServer: {
    port: 3000,
    https: true
  }
};
```

## <a name="update-the-app"></a><span data-ttu-id="1a51b-124">アプリを更新する</span><span class="sxs-lookup"><span data-stu-id="1a51b-124">Update the app</span></span>

1. <span data-ttu-id="1a51b-125">`public/index.html` ファイルを開き、`</head>` タグの直前に次の `<script>` タグを追加します。</span><span class="sxs-lookup"><span data-stu-id="1a51b-125">Open `public/index.html`, add the following `<script>` tag immediately before the `</head>` tag, and save the file.</span></span>

   ```html
   <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
   ```

2. <span data-ttu-id="1a51b-126">`src/main.js` を開き、内容を次のコードで置き換えます。</span><span class="sxs-lookup"><span data-stu-id="1a51b-126">Open the `src/main.js` file and replace its contents with the following code:</span></span>

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

3. <span data-ttu-id="1a51b-127">`src/App.vue` を開き、ファイル内容を次のコードで置き換えます。</span><span class="sxs-lookup"><span data-stu-id="1a51b-127">Open the `src/App.vue` file and replace its contents with the following code:</span></span>

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

## <a name="start-the-dev-server"></a><span data-ttu-id="1a51b-128">開発用サーバーの起動</span><span class="sxs-lookup"><span data-stu-id="1a51b-128">Start the dev server</span></span>

1. <span data-ttu-id="1a51b-129">ターミナルから、次のコマンドを実行してデベロッパー サーバーを起動します。</span><span class="sxs-lookup"><span data-stu-id="1a51b-129">From the terminal, run the following command to start the dev server.</span></span>

   ```command&nbsp;line
   npm run serve
   ```

2. <span data-ttu-id="1a51b-130">Web ブラウザーで `https://localhost:3000` (`https` に注意) に移動します。</span><span class="sxs-lookup"><span data-stu-id="1a51b-130">In a web browser, navigate to `https://localhost:3000` (notice the `https`).</span></span> <span data-ttu-id="1a51b-131">ブラウザーにサイトの証明書が信頼されていないことが示された場合は、その[証明書を信頼するようコンピューターを構成する](https://github.com/OfficeDev/generator-office/blob/fd600bbe00747e64aa5efb9846295a3f66d428aa/src/docs/ssl.md#add-certification-file-through-ie)必要があります。</span><span class="sxs-lookup"><span data-stu-id="1a51b-131">If your browser indicates that the site's certificate is not trusted, you will need to configure your computer to trust the certificate.</span></span>

3. <span data-ttu-id="1a51b-132">`https://localhost:3000` のページが空白で、証明書エラーがない場合、それは機能していることを意味します。</span><span class="sxs-lookup"><span data-stu-id="1a51b-132">When the page on `https://localhost:3000` is blank and without any certificate errors, it means that it is working.</span></span> <span data-ttu-id="1a51b-133">Vue アプリは、Office の初期化後にマウントされるため、Excel 環境内のもののみを表示します。</span><span class="sxs-lookup"><span data-stu-id="1a51b-133">The Vue App is mounted after Office is initialized, so it only shows things inside of an Excel environment.</span></span>

## <a name="try-it-out"></a><span data-ttu-id="1a51b-134">試してみる</span><span class="sxs-lookup"><span data-stu-id="1a51b-134">Try it out</span></span>

1. <span data-ttu-id="1a51b-135">アドインを実行して、Excel 内のアドインをサイドロードするのに使用するプラットフォームの手順に従います。</span><span class="sxs-lookup"><span data-stu-id="1a51b-135">Follow the instructions for the platform you'll be using to run your add-in and sideload the add-in within Excel.</span></span>

   - <span data-ttu-id="1a51b-136">Windows: [Windows で Office アドインをサイドロードする](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="1a51b-136">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
   - <span data-ttu-id="1a51b-137">Web ブラウザー: [Office on the web で Office アドインをサイドロードする](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web)</span><span class="sxs-lookup"><span data-stu-id="1a51b-137">Web browser: [Sideload Office Add-ins in Office on the web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web)</span></span>
   - <span data-ttu-id="1a51b-138">iPad および Mac: [iPad と Mac で Office アドインをサイドロードする](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="1a51b-138">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

2. <span data-ttu-id="1a51b-139">Excel で、**[ホーム]** タブを選択し、リボンの **[作業ウィンドウの表示]** ボタンをクリックして、アドインの作業ウィンドウを開きます。</span><span class="sxs-lookup"><span data-stu-id="1a51b-139">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

   ![Excel アドイン ボタン](../images/excel-quickstart-addin-2a.png)

3. <span data-ttu-id="1a51b-141">ワークシート内で任意のセルの範囲を選択します。</span><span class="sxs-lookup"><span data-stu-id="1a51b-141">Select any range of cells in the worksheet.</span></span>

4. <span data-ttu-id="1a51b-142">作業ウィンドウで、**[色の設定]** ボタンをクリックして、選択範囲の色を緑に設定します。</span><span class="sxs-lookup"><span data-stu-id="1a51b-142">In the task pane, choose the **Set color** button to set the color of the selected range to green.</span></span>

   ![Excel アドイン](../images/excel-quickstart-addin-2c.png)

## <a name="next-steps"></a><span data-ttu-id="1a51b-144">次の手順</span><span class="sxs-lookup"><span data-stu-id="1a51b-144">Next steps</span></span>

<span data-ttu-id="1a51b-145">おめでとうございます! これで Vue を使用して Excel 作業ウィンドウ アドインを作成できました。</span><span class="sxs-lookup"><span data-stu-id="1a51b-145">Congratulations, you've successfully created an Excel task pane add-in using Vue!</span></span> <span data-ttu-id="1a51b-146">次に、Excel アドインの機能の詳細について説明します。Excel アドインのチュートリアルに従って、より複雑なアドインをビルドします。</span><span class="sxs-lookup"><span data-stu-id="1a51b-146">Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="1a51b-147">Excel アドインのチュートリアル</span><span class="sxs-lookup"><span data-stu-id="1a51b-147">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.md)

## <a name="see-also"></a><span data-ttu-id="1a51b-148">関連項目</span><span class="sxs-lookup"><span data-stu-id="1a51b-148">See also</span></span>

* [<span data-ttu-id="1a51b-149">Excel アドインのチュートリアル</span><span class="sxs-lookup"><span data-stu-id="1a51b-149">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial-create-table.md)
* [<span data-ttu-id="1a51b-150">Excel JavaScript API を使用した基本的なプログラミングの概念</span><span class="sxs-lookup"><span data-stu-id="1a51b-150">Fundamental programming concepts with the Excel JavaScript API</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="1a51b-151">Excel アドインのコード サンプル</span><span class="sxs-lookup"><span data-stu-id="1a51b-151">Excel add-in code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="1a51b-152">Excel JavaScript API リファレンス</span><span class="sxs-lookup"><span data-stu-id="1a51b-152">Excel JavaScript API reference</span></span>](/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
