多くの場合、アドインは Office for Mac でキャッシュされるため、パフォーマンス上の理由があります。 通常、キャッシュはアドインを再読み込みすることでクリアされます。 同じドキュメントに複数のアドインが存在する場合、リロード時にキャッシュを自動的に消去するプロセスは、信頼できない場合があります。

作業ウィンドウアドインの [個性] メニューを使用して、キャッシュをクリアできます。
- [パーソナリティ] メニューを選択します。 [ **Web キャッシュのクリア**] を選択します。
    > [!NOTE]
    > パーソナリティメニューを表示するには、macOS バージョン10.13.6 以降を実行する必要があります。
    
    ![パーソナリティメニューの [web キャッシュのクリア] オプションのスクリーンショット。](../images/mac-clear-cache-menu.png)

`~/Library/Containers/com.Microsoft.OsfWebHost/Data/`フォルダーの内容を削除して、手動でキャッシュを消去することもできます。

> [!NOTE]
> そのフォルダーが存在しない場合は、次のフォルダーをチェックして、見つかった場合はフォルダーの内容を削除します。
>    - `~/Library/Containers/com.microsoft.{host}/Data/Library/Caches/`ここ`{host}`で、は Office ホストです (例`Excel`:)。
>    - `com.microsoft.Office365ServiceV2/Data/Caches/com.microsoft.Office365ServiceV2/`
