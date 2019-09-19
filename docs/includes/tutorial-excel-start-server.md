<span data-ttu-id="19299-101">ローカル web サーバーが既に実行されていて、アドインが既に Excel に読み込まれている場合は、手順2に進みます。</span><span class="sxs-lookup"><span data-stu-id="19299-101">If the local web server is already running and your add-in is already loaded in Excel, proceed to step 2.</span></span> <span data-ttu-id="19299-102">それ以外の場合は、ローカル web サーバーを起動して、アドインをサイドロードします。</span><span class="sxs-lookup"><span data-stu-id="19299-102">Otherwise, start the local web server and sideload your add-in:</span></span> 

- <span data-ttu-id="19299-103">Excel でアドインをテストするには、プロジェクトのルートディレクトリで次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="19299-103">To test your add-in in Excel, run the following command in the root directory of your project.</span></span> <span data-ttu-id="19299-104">これにより、ローカル web サーバーが起動 (まだ実行されていない場合) し、アドインが読み込まれた状態で Excel が開きます。</span><span class="sxs-lookup"><span data-stu-id="19299-104">This starts the local web server (if it's not already running) and opens Excel with your add-in loaded.</span></span>

    ```command&nbsp;line
    npm start
    ```

- <span data-ttu-id="19299-105">Web 上の Excel でアドインをテストするには、プロジェクトのルートディレクトリで次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="19299-105">To test your add-in in Excel on the web, run the following command in the root directory of your project.</span></span> <span data-ttu-id="19299-106">このコマンドを実行すると、ローカル web サーバーが起動します (まだ実行していない場合)。</span><span class="sxs-lookup"><span data-stu-id="19299-106">When you run this command, the local web server will start (if it's not already running).</span></span>

    ```command&nbsp;line
    npm run start:web
    ```

    <span data-ttu-id="19299-107">アドインを使用するには、web 上の Excel で新しいドキュメントを開き、「 [office のサイドロード Office アドイン](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web)」の手順に従ってアドインをサイドロードします。</span><span class="sxs-lookup"><span data-stu-id="19299-107">To use your add-in, open a new document in Excel on the web and then sideload your add-in by following the instructions in [Sideload Office Add-ins in Office on the web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).</span></span>
