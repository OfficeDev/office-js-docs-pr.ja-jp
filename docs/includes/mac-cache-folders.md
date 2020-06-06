<span data-ttu-id="ad218-101">多くの場合、アドインは Office for Mac でキャッシュされるため、パフォーマンス上の理由があります。</span><span class="sxs-lookup"><span data-stu-id="ad218-101">Add-ins are often cached in Office for Mac, for performance reasons.</span></span> <span data-ttu-id="ad218-102">通常、キャッシュはアドインを再読み込みすることでクリアされます。</span><span class="sxs-lookup"><span data-stu-id="ad218-102">Normally, the cache is cleared by reloading the add-in.</span></span> <span data-ttu-id="ad218-103">同じドキュメントに複数のアドインが存在する場合、リロード時にキャッシュを自動的に消去するプロセスは、信頼できない場合があります。</span><span class="sxs-lookup"><span data-stu-id="ad218-103">If more than one add-in exists in the same document, the process of automatically clearing the cache on reload might not be reliable.</span></span>

<span data-ttu-id="ad218-104">作業ウィンドウアドインの [個性] メニューを使用して、キャッシュをクリアできます。</span><span class="sxs-lookup"><span data-stu-id="ad218-104">You can clear the cache by using the personality menu of any task pane add-in.</span></span>
- <span data-ttu-id="ad218-105">[パーソナリティ] メニューを選択します。</span><span class="sxs-lookup"><span data-stu-id="ad218-105">Choose the personality menu.</span></span> <span data-ttu-id="ad218-106">[ **Web キャッシュのクリア**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="ad218-106">Then choose **Clear Web Cache**.</span></span>
    > [!NOTE]
    > <span data-ttu-id="ad218-107">パーソナリティメニューを表示するには、macOS バージョン10.13.6 以降を実行する必要があります。</span><span class="sxs-lookup"><span data-stu-id="ad218-107">You must run macOS version 10.13.6 or later to see the personality menu.</span></span>
    
    ![パーソナリティメニューの [web キャッシュのクリア] オプションのスクリーンショット。](../images/mac-clear-cache-menu.png)

<span data-ttu-id="ad218-109">フォルダーの内容を削除して、手動でキャッシュを消去することもでき `~/Library/Containers/com.Microsoft.OsfWebHost/Data/` ます。</span><span class="sxs-lookup"><span data-stu-id="ad218-109">You can also clear the cache manually by deleting the contents of the `~/Library/Containers/com.Microsoft.OsfWebHost/Data/` folder.</span></span>

> [!NOTE]
> <span data-ttu-id="ad218-110">そのフォルダーが存在しない場合は、次のフォルダーを確認し、見つかった場合はフォルダーの内容を削除します。</span><span class="sxs-lookup"><span data-stu-id="ad218-110">If that folder doesn't exist, check for the following folders and if found, delete the contents of the folder:</span></span>
>    - <span data-ttu-id="ad218-111">`{host}` が Office ホスト (例: `Excel`) の `~/Library/Containers/com.microsoft.{host}/Data/Library/Caches/`</span><span class="sxs-lookup"><span data-stu-id="ad218-111">`~/Library/Containers/com.microsoft.{host}/Data/Library/Caches/` where `{host}` is the Office host (e.g., `Excel`)</span></span>
>    - <span data-ttu-id="ad218-112">`{host}` が Office ホスト (例: `Excel`) の `~/Library/Containers/com.microsoft.{host}/Data/Library/Application Support/Microsoft/Office/16.0/Wef/`</span><span class="sxs-lookup"><span data-stu-id="ad218-112">`~/Library/Containers/com.microsoft.{host}/Data/Library/Application Support/Microsoft/Office/16.0/Wef/` where `{host}` is the Office host (e.g., `Excel`)</span></span>
>    - `com.microsoft.Office365ServiceV2/Data/Caches/com.microsoft.Office365ServiceV2/`
>    - `com.microsoft.Office365ServiceV2/Data/Library/Caches/com.microsoft.Office365ServiceV2/`
