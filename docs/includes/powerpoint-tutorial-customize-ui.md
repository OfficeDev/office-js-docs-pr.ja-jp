このチュートリアルの手順では、作業ウィンドウのユーザー インターフェイス (UI) をカスタマイズします。

> [!NOTE]
> このページでは、PowerPoint アドインのチュートリアルの個々の手順について説明します。 このページに検索エンジンの結果から、または直接リンクからアクセスした場合は、「[PowerPoint アドインのチュートリアル](../tutorials/powerpoint-tutorial.yml)」の紹介ページに移動し、チュートリアルを最初から始めてください。

## <a name="customize-the-task-pane-ui"></a>作業ウィンドウの UI をカスタマイズする 

1. **Home.html** ファイルで `TODO2` を次のマークアップと置き換え、ヘッダー セクションとタイトルを作業ウィンドウに追加します。 注意:

    - `ms-` で始まるスタイルは、[Office UI Fabric](../design/office-ui-fabric.md) で定義されています。これは、Office と Office 365 のユーザー エクスペリエンスを構築するための JavaScript フロント エンドのフレームワークです。 **Home.html** ファイルには、Fabric スタイル シートへの参照が含まれています。

    ```html
    <div id="content-header">
        <div class="ms-Grid ms-bgColor-neutralPrimary">
            <div class="ms-Grid-row">
                <div class="padding ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12"> <div class="ms-font-xl ms-fontColor-white ms-fontWeight-semibold">My PowerPoint add-in</div></div>
            </div>
        </div>
    </div>
    ```

2. **Home.html** ファイルにおいて、`class="footer"` で **div** を検索し、**div** 全体を削除して作業ウィンドウからフッター セクションを削除します。

## <a name="test-the-add-in"></a>アドインのテスト

1. Visual Studio を使用して、PowerPoint アドインをテストします。そのために、`F5` キーを押すか **[開始]** ボタンをクリックして、リボンに **[作業ウィンドウの表示]** アドイン ボタンが表示された PowerPoint を起動します。アドインは IIS 上でローカルにホストされます。

    ![[開始] ボタンが強調表示されている Visual Studio のスクリーンショット](../images/powerpoint-tutorial-start.png)

2. PowerPoint でリボンの **[作業ウィンドウの表示]** ボタンをクリックし、アドインの作業ウィンドウを開きます。

    ![[ホーム] リボンで [作業ウィンドウの表示] ボタンが強調表示されている Visual Studio のスクリーンショット](../images/powerpoint-tutorial-show-taskpane-button.png)

3. このとき、作業ウィンドウにはヘッダー セクションとタイトルが含まれ、フッター セクションが含まれないことがわかります。

    ![[イメージの挿入] ボタンが強調表示されている PowerPoint アドインのスクリーンショット](../images/powerpoint-tutorial-new-task-pane-ui.png)

4. Visual Studio で `Shift + F5` を押すか **[停止]** ボタンを選択してアドインを停止します。 アドインが停止すると、PowerPoint は自動的に閉じます。

    ![[停止] ボタンが強調表示されている Visual Studio のスクリーンショット](../images/powerpoint-tutorial-stop.png)

