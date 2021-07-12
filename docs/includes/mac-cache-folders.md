<span data-ttu-id="cc3cf-p101">アドインはパフォーマンス上の理由から、Office for Mac でキャッシュされることが多いです。通常、キャッシュはアドインを再読み込みすることでクリアされます。同じドキュメント内に複数のアドインが存在する場合、再読み込み時にキャッシュを自動的にクリアするプロセスは信頼できない場合があります。</span><span class="sxs-lookup"><span data-stu-id="cc3cf-p101">Add-ins are often cached in Office for Mac, for performance reasons. Normally, the cache is cleared by reloading the add-in. If more than one add-in exists in the same document, the process of automatically clearing the cache on reload might not be reliable.</span></span>

<span data-ttu-id="cc3cf-104">作業ウィンドウ アドインの [パーソナリティ] メニューを使用してキャッシュをクリアすることができます。</span><span class="sxs-lookup"><span data-stu-id="cc3cf-104">You can clear the cache by using the personality menu of any task pane add-in.</span></span>
- <span data-ttu-id="cc3cf-105">[パーソナリティ] メニューを選択します。</span><span class="sxs-lookup"><span data-stu-id="cc3cf-105">Choose the personality menu.</span></span> <span data-ttu-id="cc3cf-106">次に、**[Web キャッシュのクリア]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="cc3cf-106">Then choose **Clear Web Cache**.</span></span>
    > [!NOTE]
    > <span data-ttu-id="cc3cf-107">[パーソナリティ] メニューを表示するには、macOS のバージョン 10.13.6 以降を実行する必要があります。</span><span class="sxs-lookup"><span data-stu-id="cc3cf-107">You must run macOS version 10.13.6 or later to see the personality menu.</span></span>

    ![[パーソナリティ] メニューの [Web キャッシュのクリア] オプションのスクリーン ショット。](../images/mac-clear-cache-menu.png)

<span data-ttu-id="cc3cf-109">`~/Library/Containers/com.Microsoft.OsfWebHost/Data/` フォルダーのコンテンツを削除することによってキャッシュを手動でクリアすることもできます。</span><span class="sxs-lookup"><span data-stu-id="cc3cf-109">You can also clear the cache manually by deleting the contents of the `~/Library/Containers/com.Microsoft.OsfWebHost/Data/` folder.</span></span>

> [!NOTE]
> <span data-ttu-id="cc3cf-110">そのフォルダーが存在しない場合には次のフォルダーを確認し、見つかった場合はフォルダーのコンテンツを削除します。</span><span class="sxs-lookup"><span data-stu-id="cc3cf-110">If that folder doesn't exist, check for the following folders and if found, delete the contents of the folder.</span></span>
>    - <span data-ttu-id="cc3cf-111">`{host}` が Office アプリケーション (例: `Excel`) である `~/Library/Containers/com.microsoft.{host}/Data/Library/Caches/`</span><span class="sxs-lookup"><span data-stu-id="cc3cf-111">`~/Library/Containers/com.microsoft.{host}/Data/Library/Caches/` where `{host}` is the Office application (e.g., `Excel`)</span></span>
>    - <span data-ttu-id="cc3cf-112">`{host}` が Office アプリケーション (例: `Excel`) である `~/Library/Containers/com.microsoft.{host}/Data/Library/Application Support/Microsoft/Office/16.0/Wef/`</span><span class="sxs-lookup"><span data-stu-id="cc3cf-112">`~/Library/Containers/com.microsoft.{host}/Data/Library/Application Support/Microsoft/Office/16.0/Wef/` where `{host}` is the Office application (e.g., `Excel`)</span></span>
>    - `~/Library/Containers/com.microsoft.Office365ServiceV2/Data/Caches/com.microsoft.Office365ServiceV2/`
>    - `~/Library/Containers/com.microsoft.Office365ServiceV2/Data/Library/Caches/com.microsoft.Office365ServiceV2/`
