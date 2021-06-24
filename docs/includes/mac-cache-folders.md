アドインはパフォーマンス上の理由から、Office for Mac でキャッシュされることが多いです。通常、キャッシュはアドインを再読み込みすることでクリアされます。同じドキュメント内に複数のアドインが存在する場合、再読み込み時にキャッシュを自動的にクリアするプロセスは信頼できない場合があります。

作業ウィンドウ アドインの [パーソナリティ] メニューを使用してキャッシュをクリアすることができます。
- [パーソナリティ] メニューを選択します。 次に、**[Web キャッシュのクリア]** を選択します。
    > [!NOTE]
    > [パーソナリティ] メニューを表示するには、macOS のバージョン 10.13.6 以降を実行する必要があります。

    ![[パーソナリティ] メニューの [Web キャッシュのクリア] オプションのスクリーン ショット。](../images/mac-clear-cache-menu.png)

`~/Library/Containers/com.Microsoft.OsfWebHost/Data/` フォルダーのコンテンツを削除することによってキャッシュを手動でクリアすることもできます。

> [!NOTE]
> そのフォルダーが存在しない場合には次のフォルダーを確認し、見つかった場合はフォルダーの内容を削除します。
>    - `{host}` が Office アプリケーション (例: `Excel`) である `~/Library/Containers/com.microsoft.{host}/Data/Library/Caches/`
>    - `{host}` が Office アプリケーション (例: `Excel`) である `~/Library/Containers/com.microsoft.{host}/Data/Library/Application Support/Microsoft/Office/16.0/Wef/`
>    - `~/Library/Containers/com.microsoft.Office365ServiceV2/Data/Caches/com.microsoft.Office365ServiceV2/`
>    - `~/Library/Containers/com.microsoft.Office365ServiceV2/Data/Library/Caches/com.microsoft.Office365ServiceV2/`
