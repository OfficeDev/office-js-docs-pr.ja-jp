<span data-ttu-id="b9ecd-p101">Office アドインは、Web アプリケーションとマニフェスト ファイルで構成されます。Web アプリケーションはアドインのユーザー インターフェイスと機能を定義しますが、マニフェストは Web アプリケーションの場所を指定し、アドインの設定と機能を定義します。</span><span class="sxs-lookup"><span data-stu-id="b9ecd-p101">An Office Add-in consists of a web application and a manifest file. The web application defines the add-in's user interface and functionality, while the manifest specifies the location of the web application and defines settings and capabilities of the add-in.</span></span> 

<span data-ttu-id="b9ecd-103">アドインの開発中に、ローカル Web サーバーでアドインを実行できます (`localhost`)。ただし、他のユーザーがアクセスできるように公開する準備ができたら、Web サーバーまたは Web ホスティング サービス (Microsoft Azure など) に Web アプリケーションを展開し、マニフェストを更新して展開されたアプリケーションの URL を指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="b9ecd-103">While you're developing your add-in, you can run the add-in on your local web server (`localhost`), but when you're ready to publish it for other users to access, you'll need to deploy the web application to a web server or web hosting service (for example, Microsoft Azure) and update the manifest to specify the URL of the deployed application.</span></span> 

<span data-ttu-id="b9ecd-104">アドインが希望どおりに機能し、他のユーザーがアクセスできるように公開する準備ができたら、次の手順を実行します。</span><span class="sxs-lookup"><span data-stu-id="b9ecd-104">When your add-in is working as desired and you're ready to publish it for other users to access, complete the following steps.</span></span>

1. <span data-ttu-id="b9ecd-105">コマンド ラインから、アドイン プロジェクトのルート ディレクトリで、次のコマンドを実行して、運用展開用のすべてのファイルを準備します。</span><span class="sxs-lookup"><span data-stu-id="b9ecd-105">From the command line, in the root directory of your add-in project, run the following command to prepare all files for production deployment.</span></span>

    ```command&nbsp;line
    npm run build
    ```

    <span data-ttu-id="b9ecd-106">ビルドが完了すると、アドイン プロジェクトのルート ディレクトリにある **dist** フォルダーに、以降の手順で展開するファイルが含まれます。</span><span class="sxs-lookup"><span data-stu-id="b9ecd-106">When the build completes, the **dist** folder in the root directory of your add-in project will contain the files that you'll deploy in subsequent steps.</span></span>

2. <span data-ttu-id="b9ecd-107">**dist** フォルダーの内容を、アドインをホストする Web サーバーにアップロードします。</span><span class="sxs-lookup"><span data-stu-id="b9ecd-107">Upload the contents of the **dist** folder to the web server that'll host your add-in.</span></span> <span data-ttu-id="b9ecd-108">任意の種類の Web サーバーまたは Web ホスティング サービスを使用して、アドインをホストできます。</span><span class="sxs-lookup"><span data-stu-id="b9ecd-108">You can use any type of web server or web hosting service to host your add-in.</span></span>

3. <span data-ttu-id="b9ecd-109">VS Code で、プロジェクトのルート ディレクトリにあるアドインのマニフェスト ファイルを開きます (`manifest.xml`)。</span><span class="sxs-lookup"><span data-stu-id="b9ecd-109">In VS Code, open the add-in's manifest file, located in the root directory of the project (`manifest.xml`).</span></span> <span data-ttu-id="b9ecd-110">`https://localhost:3000` のすべての出現回数を、前の手順で Web サーバーに展開した Web アプリケーションの URL に置き換えます。</span><span class="sxs-lookup"><span data-stu-id="b9ecd-110">Replace all occurrences of `https://localhost:3000` with the URL of the web application that you deployed to a web server in the previous step.</span></span>

4. <span data-ttu-id="b9ecd-111">[Office アドインを展開](../publish/publish.md)するために使用する方法を選択し、指示に従ってマニフェスト ファイルを公開します。</span><span class="sxs-lookup"><span data-stu-id="b9ecd-111">Choose the method you'd like to use to [deploy your Office Add-in](../publish/publish.md), and follow the instructions to publish the manifest file.</span></span>
