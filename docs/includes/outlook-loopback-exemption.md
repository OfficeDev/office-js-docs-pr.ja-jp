> [!NOTE]
> Outlook Windows: localhost からアドインを実行している場合は、「申し訳ありませんが、{*your-add-in-name-here}* にアクセスできませんでした。」というエラーが表示されます。 ネットワーク接続が確立されている必要があります。 問題が解決しない場合は、後でもう一度お試しください。ループバックの除外を有効にする必要がある場合があります。
>
> 1. Outlook を終了します。
> 1. タスク マネージャー **を開** き、 **タスクmsoadfsb.exeが** 実行されていないか確認します。
> 1. 管理者特権 [のプロンプトで](/previous-versions/windows/apps/hh780593(v=win.10)?redirectedfrom=MSDN) ループバックの除外を設定します。
>     - `https://localhost`ポート 3000 (既定の構成) を使用している場合は、次のコマンドを実行します。
>
>        ```command&nbsp;line
>        call %SystemRoot%\System32\CheckNetIsolation.exe LoopbackExempt -a -n=1_https___localhost_300004ACA5EC-D79A-43EA-AB47-E50E47DD96FC
>        ```
>     - `http://localhost`ポート 3000 を使用している場合は、次のコマンドを実行します。
>
>        ```command&nbsp;line
>        call %SystemRoot%\System32\CheckNetIsolation.exe LoopbackExempt -a -n=1_http___localhost_300004ACA5EC-D79A-43EA-AB47-E50E47DD96FC
>        ```
>
>      **注**: 既定のポート 3000 を使用していない場合は、コマンドで実際のポート番号に置き換える必要があります。
> 1. Outlook を再起動します。
