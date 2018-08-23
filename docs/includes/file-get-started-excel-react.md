# <a name="build-an-excel-add-in-using-react"></a>React を使用して Excel のアドインを作成する

この記事では、React と Excel の JavaScript API を使用して Excel アドインを構築する手順を説明します。

## <a name="environment"></a>環境

- **Office Desktop**最新バージョンのOfficeがインストールされていることを確認してください。 アドインコマンドにはビルド 16.0.6769.0000 以上が必要です (**16.0.6868.0000** 推奨)。 [Office アプリケーションの最新バージョンをインストールする](http://aka.ms/latestoffice)方法。 
 
- **Office Online**：追加設定はありません。 Office Online の職場/学校アカウント用コマンドのサポートはプレビューになっています。

## <a name="prerequisites"></a>前提条件

- [Node.js](https://nodejs.org)

- [Yeoman](https://github.com/yeoman/yo) の最新バージョンと [Office アドイン用の Yeoman ジェネレーター](https://github.com/OfficeDev/generator-office)をグローバルにインストールします。
    ```bash
    npm install -g yo generator-office
    ```

### <a name="create-the-web-app"></a>Web アプリを作成する

1. ローカル ドライブにフォルダーを作成し、**my-addin** という名前を付けます。 ここにアプリのファイルを作成します。

2. アプリ フォルダーに移動します。

    ```bash
    cd my-addin
    ```

3. Yeoman ジェネレーター使用して、アドインのマニフェスト ファイルを生成します。 次のコマンドを実行し、以下のスクリーンショットに示すとおり、プロンプトに応答します。

    ```bash
    yo office
    ```

    - **Choose a project type:​ (プロジェクト タイプを選択してください)** `Office Add-in project using React framework`
    - **What would you want to name your add-in?: (アドインの名前を何にしますか)** `My Office Add-in`
    - **Which Office client application would you like to support?: (どの Office クライアント アプリケーションをサポートしますか)** `Excel`

    ![Yeoman ジェネレーター](../images/yo-office-excel-react.png)
    
    ウィザードが完了すると、ジェネレーターはプロジェクトを作成し、サポートする Node コンポーネントをインストールします。

4.   **Src/components/App.tsx** を開き、コメントの「塗りつぶしの色を更新する」を検索し、塗りつぶしの色を '青' から'黄' に変更してから保存します。 

    ```js
    range.format.fill.color = 'blue'

    ```

5. **src/components/App.tsx** 内の `render` 関数の `return` ブロックで、`<Herolist>` を次のコードに更新し、ファイルを保存します。 

    ```js
      <HeroList message='Discover what My Office Add-in can do for you today!' items={this.state.listItems}>
        <p className='ms-font-l'>Choose the button below to set the color of the selected range to blue. <b>Set color</b>.</p>
        <Button className='ms-welcome__action' buttonType={ButtonType.hero} iconProps={{ iconName: 'ChevronRight' }} onClick={this.click}>Run</Button>
    </HeroList>
    ```

6. 「[自己署名証明書を信頼されたルート証明書として追加する](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md)」の手順を実行して、開発用コンピューターのオペレーティング システムの証明書を信頼します。

7. アドインをサイドロードすると Excel に表示されます。 ターミナルで次のコマンドを実行します。 
    
    ```bash
    npm run sideload
    ```

## <a name="try-it-out"></a>お試しください

1. ターミナルから、次のコマンドを実行してデベロッパー サーバーを起動します。

    Windows
    ```bash
    npm start
    ```

2. Excel で、**[ホーム]** タブを選択し、リボンの **[作業ウィンドウの表示]** ボタンをクリックして、アドインの作業ウィンドウを開きます。

    ![Excel アドイン ボタン](../images/excel-quickstart-addin-2b.png)

3. ワークシート内で任意のセルの範囲を選択します。

4. 作業ウィンドウで **[色の設定]** ボタンをクリックし、選択範囲の色を青に設定します。

    ![Excel アドイン](../images/excel-quickstart-addin-2c.png)

## <a name="next-steps"></a>次の手順

これで完了です。React を使用して Excel アドインが正常に作成されました。次に、Excel アドインの機能の詳細について説明します。Excel アドインのチュートリアルに従って、より複雑なアドインをビルドします。

> [!div class="nextstepaction"]
> [Excel アドインのチュートリアル](../tutorials/excel-tutorial.yml)

## <a name="see-also"></a>関連項目

* [Excel アドインのチュートリアル](../tutorials/excel-tutorial-create-table.md)
* [Excel JavaScript API の中心概念](../excel/excel-add-ins-core-concepts.md)
* [Excel アドインのコード サンプル](http://dev.office.com/code-samples#?filters=excel,office%20add-ins)
* [Excel JavaScript API リファレンス](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview)
