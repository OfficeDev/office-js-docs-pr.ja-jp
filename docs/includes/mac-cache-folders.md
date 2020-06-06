多くの場合、アドインは Office for Mac でキャッシュされるため、パフォーマンス上の理由があります。 通常、キャッシュはアドインを再読み込みすることでクリアされます。 同じドキュメントに複数のアドインが存在する場合、リロード時にキャッシュを自動的に消去するプロセスは、信頼できない場合があります。

作業ウィンドウアドインの [個性] メニューを使用して、キャッシュをクリアできます。
- [パーソナリティ] メニューを選択します。 [ **Web キャッシュのクリア**] を選択します。
    > [!NOTE]
    > パーソナリティメニューを表示するには、macOS バージョン10.13.6 以降を実行する必要があります。
    
    ![パーソナリティメニューの [web キャッシュのクリア] オプションのスクリーンショット。](../images/mac-clear-cache-menu.png)

フォルダーの内容を削除して、手動でキャッシュを消去することもでき `~/Library/Containers/com.Microsoft.OsfWebHost/Data/` ます。

> [!NOTE]
> そのフォルダーが存在しない場合は、次のフォルダーを確認し、見つかった場合はフォルダーの内容を削除します。
>    - `{host}` が Office ホスト (例: `Excel`) の `~/Library/Containers/com.microsoft.{host}/Data/Library/Caches/`
>    - `{host}` が Office ホスト (例: `Excel`) の `~/Library/Containers/com.microsoft.{host}/Data/Library/Application Support/Microsoft/Office/16.0/Wef/`
>    - `com.microsoft.Office365ServiceV2/Data/Caches/com.microsoft.Office365ServiceV2/`
>    - `com.microsoft.Office365ServiceV2/Data/Library/Caches/com.microsoft.Office365ServiceV2/`
