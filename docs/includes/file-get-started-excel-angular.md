# <a name="build-an-excel-add-in-using-angular"></a><span data-ttu-id="63dc5-101">Angular を使用して Excel のアドインを作成する</span><span class="sxs-lookup"><span data-stu-id="63dc5-101">Build an Excel add-in using Angular</span></span>

<span data-ttu-id="63dc5-102">この記事では、Angular と Excel の JavaScript API を使用して Excel アドインを構築する手順について説明します。</span><span class="sxs-lookup"><span data-stu-id="63dc5-102">In this article, you'll walk you through the process of building an Excel add-in using Angular and the Excel JavaScript API.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="63dc5-103">前提条件</span><span class="sxs-lookup"><span data-stu-id="63dc5-103">Prerequisites</span></span>

- <span data-ttu-id="63dc5-104">既に [Angular CLI の必須コンポーネント](https://github.com/angular/angular-cli#prerequisites)がインストールされているかを確認し、足りない場合は必須コンポーネントをインストールします。</span><span class="sxs-lookup"><span data-stu-id="63dc5-104">Check whether you already have the [Angular CLI prerequisites](https://github.com/angular/angular-cli#prerequisites) and install any prerequistes that you are missing.</span></span>

- <span data-ttu-id="63dc5-105">[Angular CLI](https://github.com/angular/angular-cli) をグローバルにインストールします。</span><span class="sxs-lookup"><span data-stu-id="63dc5-105">Install the [Angular CLI](https://github.com/angular/angular-cli) globally.</span></span> 

    ```bash
    npm install -g @angular/cli
    ```

- <span data-ttu-id="63dc5-106">[Yeoman](https://github.com/yeoman/yo) の最新バージョンと [Office アドイン用の Yeoman ジェネレーター](https://github.com/OfficeDev/generator-office)をグローバルにインストールします。</span><span class="sxs-lookup"><span data-stu-id="63dc5-106">Install the latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globally.</span></span>

    ```bash
    npm install -g yo generator-office
    ```

## <a name="generate-a-new-angular-app"></a><span data-ttu-id="63dc5-107">新しい Angular アプリを生成する</span><span class="sxs-lookup"><span data-stu-id="63dc5-107">Generate a new Angular app</span></span>

<span data-ttu-id="63dc5-p101">Angular CLI を使用して Angular アプリを生成します。端末では、次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="63dc5-p101">Use the Angular CLI to generate your Angular app. From the terminal, run the following command:</span></span>

```bash
ng new my-addin
```

## <a name="generate-the-manifest-file"></a><span data-ttu-id="63dc5-110">マニフェスト ファイルを生成する</span><span class="sxs-lookup"><span data-stu-id="63dc5-110">Generate the manifest file</span></span>

<span data-ttu-id="63dc5-111">アドインのマニフェスト ファイルは、その設定と機能を定義します。</span><span class="sxs-lookup"><span data-stu-id="63dc5-111">An add-in's manifest file defines its settings and capabilities.</span></span>

1. <span data-ttu-id="63dc5-112">アプリ フォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="63dc5-112">Navigate to your app folder.</span></span>

    ```bash
    cd my-addin
    ```

2. <span data-ttu-id="63dc5-p102">Yeoman ジェネレーター使用して、アドインのマニフェスト ファイルを生成します。次のコマンドを実行し、以下に示すとおりにプロンプトに応答します。</span><span class="sxs-lookup"><span data-stu-id="63dc5-p102">Use the Yeoman generator to generate the manifest file for your add-in. Run the following command and then answer the prompts as shown below.</span></span>

    ```bash
    yo office 
    ```

    - <span data-ttu-id="63dc5-115">**Choose a project type:​ (プロジェクト タイプを選択してください:)** `Office Add-in containing the manifest only`</span><span class="sxs-lookup"><span data-stu-id="63dc5-115">**Choose a project type:** `Office Add-in containing the manifest only`</span></span>
    - <span data-ttu-id="63dc5-116">**What would you want to name your add-in?: (アドインの名前を何にしますか)** `My Office Add-in`</span><span class="sxs-lookup"><span data-stu-id="63dc5-116">**What do you want to name your add-in?:** `My Office Add-in`</span></span>
    - <span data-ttu-id="63dc5-117">**どの Office クライアント アプリケーションをサポートしますか？** `Excel`</span><span class="sxs-lookup"><span data-stu-id="63dc5-117">**Which Office client application would you like to support?:** `Excel`</span></span>

    <span data-ttu-id="63dc5-118">ウィザードを完了すると、マニフェスト ファイルとリソース ファイルを使用してプロジェクトをビルドできます。</span><span class="sxs-lookup"><span data-stu-id="63dc5-118">After you complete the wizard, a manifest file and resource file are available for you to build your project.</span></span>

    ![Yeoman ジェネレーター](../images/yo-office.png)
    
    > [!NOTE]
    > <span data-ttu-id="63dc5-120">**package.json** を上書きするメッセージが表示された場合は、**No** (上書きしない) と応答します。</span><span class="sxs-lookup"><span data-stu-id="63dc5-120">If you're prompted to overwrite **package.json**, answer **No** (do not overwrite).</span></span>

## <a name="secure-the-app"></a><span data-ttu-id="63dc5-121">アプリをセキュリティ保護する</span><span class="sxs-lookup"><span data-stu-id="63dc5-121">Secure the app</span></span>

[!include[HTTPS guidance](../includes/https-guidance.md)]

<span data-ttu-id="63dc5-p103">このクイック スタートでは、証明書を使用できますが、 **Yeoman Office アドイン用のジェネレーター** が用意されています。ジェネレーターをグローバルに (このクイック スタートの **前提条件** の一部) としてインストールして、グローバルから証明書をコピーする必要がありますだけで、インストールの場所、アプリケーション フォルダーにするとします。次の手順では、このプロセスを完了する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="63dc5-p103">For this quick start, you can use the certificates that the **Yeoman generator for Office Add-ins** provides. You've already installed the generator globally (as part of the **Prerequisites** for this quick start), so you'll just need to copy the certificates from the global install location into your app folder. The following steps describe how to complete this process.</span></span>

1. <span data-ttu-id="63dc5-125">端末から次のコマンドを実行し、グローバル **npm** ライブラリがインストールされているフォルダーを識別します。</span><span class="sxs-lookup"><span data-stu-id="63dc5-125">From the terminal, run the following command to identify the folder where global **npm** libraries are installed:</span></span>

    ```bash 
    npm list -g 
    ``` 
    
    > [!TIP]    
    > <span data-ttu-id="63dc5-126">このコマンドで生成された出力の最初の行は、グローバル **npm** ライブラリがインストールされているフォルダーを示します。</span><span class="sxs-lookup"><span data-stu-id="63dc5-126">The first line of output that's generated by this command specifies the folder where global **npm** libraries are installed.</span></span>          
    
2. <span data-ttu-id="63dc5-p104">ファイル エクスプローラーを使用して`{global libraries folder}/node_modules/generator-office/generators/app/templates/js/base`フォルダーに移動します。その場所から`certs`フォルダーをクリップボードにコピーします。</span><span class="sxs-lookup"><span data-stu-id="63dc5-p104">Using File Explorer, navigate to the `{global libraries folder}/node_modules/generator-office/generators/app/templates/js/base` folder. From that location, copy the `certs` folder to your clipboard.</span></span>

3. <span data-ttu-id="63dc5-129">前のセクションの手順 1 で作成した Angular アプリのルート フォルダーに移動して、クリップボードからそのフォルダーに `certs` フォルダーを貼り付けます。</span><span class="sxs-lookup"><span data-stu-id="63dc5-129">Navigate to the root folder of the Angular app that you created in step 1 of the previous section, and paste the `certs` folder from your clipboard into that folder.</span></span>

## <a name="update-the-app"></a><span data-ttu-id="63dc5-130">アプリを更新する</span><span class="sxs-lookup"><span data-stu-id="63dc5-130">Update the app</span></span>

1. <span data-ttu-id="63dc5-p105">コード エディターでプロジェクトのルートに **package.json** を開きます。変更、 `start` サーバーが SSL およびポート 3000 を使用して実行し、ファイルを保存するを指定するためのスクリプトです。</span><span class="sxs-lookup"><span data-stu-id="63dc5-p105">In your code editor, open **package.json** in the root of the project. Modify the `start` script to specify that the server should run using SSL and port 3000, and save the file.</span></span>

    ```json
    "start": "ng serve --ssl true --port 3000"
    ```

2. <span data-ttu-id="63dc5-p106">プロジェクトのルートに **.angular cli.json** を開きます。証明書ファイルの場所を指定する **既定の設定** オブジェクトを変更し、ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="63dc5-p106">Open **.angular-cli.json** in the root of the project. Modify the **defaults** object to specify the location of the certificate files, and save the file.</span></span>

    ```json
    "defaults": {
      "styleExt": "css",
      "component": {},
      "serve": {
        "sslKey": "certs/server.key",
        "sslCert": "certs/server.crt"
      }
    }
    ```

3. <span data-ttu-id="63dc5-135">**src/index.html** を開き、`</head>` タグの直前に次の `<script>` タグを追加して、ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="63dc5-135">Open **src/index.html**, add the following `<script>` tag immediately before the `</head>` tag, and save the file.</span></span>

    ```html
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
    ```

4. <span data-ttu-id="63dc5-136">**src/main.ts** を開き、`platformBrowserDynamic().bootstrapModule(AppModule).catch(err => console.log(err));` を次のコードで置き換えて、ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="63dc5-136">Open **src/main.ts**, replace `platformBrowserDynamic().bootstrapModule(AppModule).catch(err => console.log(err));` with the following code, and save the file.</span></span> 

    ```typescript 
    declare const Office: any;

    Office.initialize = () => {
    platformBrowserDynamic().bootstrapModule(AppModule)
        .catch(err => console.log(err));
    };
    ```

5. <span data-ttu-id="63dc5-137">**src/polyfills.ts** を開き、存在する他のすべての `import` ステートメントの上に次のコード行を追加して、ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="63dc5-137">Open **src/polyfills.ts**, add the following line of code above all other existing `import` statements, and save the file.</span></span>

    ```typescript
    import 'core-js/client/shim';
    ```

6. <span data-ttu-id="63dc5-138">**src/polyfills.ts** で、次の行のコメントを解除してファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="63dc5-138">In **src/polyfills.ts**, uncomment the following lines, and save the file.</span></span>

    ```typescript
    import 'core-js/es6/symbol';
    import 'core-js/es6/object';
    import 'core-js/es6/function';
    import 'core-js/es6/parse-int';
    import 'core-js/es6/parse-float';
    import 'core-js/es6/number';
    import 'core-js/es6/math';
    import 'core-js/es6/string';
    import 'core-js/es6/date';
    import 'core-js/es6/array';
    import 'core-js/es6/regexp';
    import 'core-js/es6/map';
    import 'core-js/es6/weak-map';
    import 'core-js/es6/set';
    ```

7. <span data-ttu-id="63dc5-139">**src/app/app.component.html** を開き、ファイルのコンテンツを次の HTML で置き換えて、ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="63dc5-139">Open **src/app/app.component.html**, replace file contents with the following HTML, and save the file.</span></span> 

    ```html
    <div id="content-header">
        <div class="padding">
            <h1>Welcome</h1>
        </div>
    </div>
    <div id="content-main">
        <div class="padding">
            <p>Choose the button below to set the color of the selected range to green.</p>
            <br />
            <h3>Try it out</h3>
            <button (click)="onSetColor()">Set color</button>
        </div>
    </div>
    ```

8. <span data-ttu-id="63dc5-140">**src/app/app.component.css** を開き、ファイルのコンテンツを次の CSS コードで置き換えて、ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="63dc5-140">Open **src/app/app.component.css**, replace file contents with the following CSS code, and save the file.</span></span>

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

9. <span data-ttu-id="63dc5-141">**src/app/app.component.ts** を開き、ファイルのコンテンツを次のコードで置き換えて、ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="63dc5-141">Open **src/app/app.component.ts**, replace file contents with the following code, and save the file.</span></span> 

    ```typescript
    import { Component } from '@angular/core';

    declare const Excel: any;

    @Component({
    selector: 'app-root',
    templateUrl: './app.component.html',
    styleUrls: ['./app.component.css']
    })
    export class AppComponent {
    onSetColor() {
        Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.format.fill.color = 'green';
        await context.sync();
        });
    }
    }
    ```

## <a name="start-the-dev-server"></a><span data-ttu-id="63dc5-142">開発用サーバーを起動する</span><span class="sxs-lookup"><span data-stu-id="63dc5-142">Start the dev server</span></span>

1. <span data-ttu-id="63dc5-143">ターミナルから、次のコマンドを実行してデベロッパー サーバーを起動します。</span><span class="sxs-lookup"><span data-stu-id="63dc5-143">From the terminal, run the following command to start the dev server.</span></span>

    ```bash
    npm run start
    ```

2. <span data-ttu-id="63dc5-p107">Web ブラウザーで `https://localhost:3000` に移動します。ブラウザーにサイトの証明書が信頼されていないことが示された場合は、その証明書を信頼された証明書として追加する必要があります。詳細については、「[自己署名証明書を信頼されたルート証明書として追加する](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="63dc5-p107">In a web browser, navigate to `https://localhost:3000`. If your browser indicates that the site's certificate is not trusted, you will need to add the certificate as a trusted certificate. See [Adding Self-Signed Certificates as Trusted Root Certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) for details.</span></span>

    > [!NOTE]
    > <span data-ttu-id="63dc5-p108">「[自己署名証明書を信頼されたルート証明書として追加する](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md)」で説明されているプロセスを完了した後でも、Chrome (Web ブラウザー) は、サイトの証明書が信頼されていないことを引き続き示すことがあります。Chrome でこの警告を無視して、Internet Explorer または Microsoft Edge のいずれかで`https://localhost:3000` に移動して証明書が信頼できるかを確認することができます。</span><span class="sxs-lookup"><span data-stu-id="63dc5-p108">Chrome (web browser) may continue to indicate the the site's certificate is not trusted, even after you have completed the process described in [Adding Self-Signed Certificates as Trusted Root Certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md). You can disregard this warning in Chrome and can verify that the certificate is trusted by navigating to `https://localhost:3000` in either Internet Explorer or Microsoft Edge.</span></span> 

3. <span data-ttu-id="63dc5-149">証明書エラーなしにブラウザにアドイン ページが読み込まれたら、アドインをテストする準備ができています。</span><span class="sxs-lookup"><span data-stu-id="63dc5-149">After your browser loads the add-in page without any certificate errors, you're ready test your add-in.</span></span> 

## <a name="try-it-out"></a><span data-ttu-id="63dc5-150">お試しください</span><span class="sxs-lookup"><span data-stu-id="63dc5-150">Try it out</span></span>

1. <span data-ttu-id="63dc5-151">アドインを実行して、Excel 内のアドインをサイドロードするために使用するプラットフォームの手順に従います。</span><span class="sxs-lookup"><span data-stu-id="63dc5-151">Follow the instructions for the platform you'll be using to run your add-in and sideload the add-in within Excel.</span></span>

    - <span data-ttu-id="63dc5-152">Windows: [Windows で Office アドインをサイドロードする](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="63dc5-152">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
    - <span data-ttu-id="63dc5-153">Excel Online:[Office Online で Office アドインをサイドロードする](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span><span class="sxs-lookup"><span data-stu-id="63dc5-153">Excel Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span></span>
    - <span data-ttu-id="63dc5-154">iPad および Mac: [iPad と Mac で Office アドインをサイドロードする](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="63dc5-154">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

   
2. <span data-ttu-id="63dc5-155">Excel で、**[ホーム]** タブを選択し、リボンの **[作業ウィンドウの表示]** ボタンをクリックして、アドインの作業ウィンドウを開きます。</span><span class="sxs-lookup"><span data-stu-id="63dc5-155">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Excel アドイン ボタン](../images/excel-quickstart-addin-2a.png)

3. <span data-ttu-id="63dc5-157">ワークシート内で任意のセル範囲を選択します。</span><span class="sxs-lookup"><span data-stu-id="63dc5-157">Select any range of cells in the worksheet.</span></span>

4. <span data-ttu-id="63dc5-158">作業ウィンドウで、**[色の設定]** ボタンをクリックして、選択範囲の色を緑に設定します。</span><span class="sxs-lookup"><span data-stu-id="63dc5-158">In the task pane, choose the **Set color** button to set the color of the selected range to green.</span></span>

    ![Excel アドイン](../images/excel-quickstart-addin-2c.png)

## <a name="next-steps"></a><span data-ttu-id="63dc5-160">次の手順</span><span class="sxs-lookup"><span data-stu-id="63dc5-160">Next steps</span></span>

<span data-ttu-id="63dc5-p109">これで完了です。Angular を使用して Excel アドインが正常に作成されました。次に、Excel アドインの機能の詳細について説明します。Excel アドインのチュートリアルに従って、より複雑なアドインをビルドします。</span><span class="sxs-lookup"><span data-stu-id="63dc5-p109">Congratulations, you've successfully created an Excel add-in using Angular! Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="63dc5-163">Excel アドインのチュートリアル</span><span class="sxs-lookup"><span data-stu-id="63dc5-163">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.yml)

## <a name="see-also"></a><span data-ttu-id="63dc5-164">関連項目</span><span class="sxs-lookup"><span data-stu-id="63dc5-164">See also</span></span>

* [<span data-ttu-id="63dc5-165">Excel アドインのチュートリアル</span><span class="sxs-lookup"><span data-stu-id="63dc5-165">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial-create-table.md)
* [<span data-ttu-id="63dc5-166">Excel の JavaScript API を使用した基本的なプログラミングの概念</span><span class="sxs-lookup"><span data-stu-id="63dc5-166">Fundamental programming concepts with the Excel JavaScript API</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="63dc5-167">Excel アドインのコード サンプル</span><span class="sxs-lookup"><span data-stu-id="63dc5-167">Excel add-in code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="63dc5-168">Excel JavaScript API リファレンス</span><span class="sxs-lookup"><span data-stu-id="63dc5-168">Excel JavaScript API reference</span></span>](https://docs.microsoft.com/javascript/office/overview/excel-add-ins-reference-overview?view=office-js)
