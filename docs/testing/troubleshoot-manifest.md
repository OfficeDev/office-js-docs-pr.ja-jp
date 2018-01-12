# <a name="validate-and-troubleshoot-issues-with-your-manifest"></a>マニフェストの問題を検証し、トラブルシューティングする

これらのメソッドを使って、マニフェストの問題を検証し、トラブルシューティングします。 

- [Office アドイン検証ツールを使用して Office アドイン マニフェストを検証する](validate-the-office-add-ins-manifest-against-validator)   
- [XML スキーマに対して、Office アドイン マニフェストを検証する](validate-the-office-add-ins-manifest-against-the-xml-schema)
- [ランタイムのログを使用して Office アドインのマニフェストをデバッグする](use-runtime-logging-to-debug-the-manifest-for-your-office-add-in)

## <a name="validate-your-manifest-with-the-office-add-in-validator"></a>Office アドイン検証ツールを使用してマニフェストを検証する
Office アドインを記述するマニフェスト ファイルが正確かつ完全であることを確認するために、[Office アドイン検証ツール](https://github.com/OfficeDev/office-addin-validator)に対してマニフェスト ファイルを検証します。

Office アドイン検証ツールを使用してマニフェストを検証するには、次のようにします。

1. [Node.js](https://nodejs.org/download/) をインストールします。 
2. 管理者としてコマンド プロンプト / terminal を開き、次のコマンドを使用して Office アドイン検証ツールとその依存関係をインストールします。

    ```
    npm install -g office-addin-validator
    ```
    
    > **注:**Yo Office が既にインストールされている場合、最新のバージョンにアップグレードすると、検証ツールが依存関係としてインストールされます。

3. マニフェストを検証するには、次のコマンドを実行します。MANIFEST.XML をマニフェスト XML ファイルへのパスに置き換えます。

    ```
    validate-office-addin MANIFEST.XML
    ```


## <a name="validate-your-manifest-against-the-xml-schema"></a>XML スキーマと比較してマニフェストを検証する

マニフェスト ファイルが正しいスキーマに従っていることを確認するために、[XML スキーマ定義 (XSD)](https://github.com/OfficeDev/office-js-docs/tree/master/docs/overview/schemas) ファイルと比較してマニフェスト ファイルを検証します。XML スキーマの検証ツールを使用して、この検証を実行できます。 

コマンド ライン XML スキーマ検証ツールを使用してマニフェストを検証するには、次のようにします。

1.  [tar](https://www.gnu.org/software/tar/) および [libxml](http://xmlsoft.org/FAQ.html) をまだインストールしていない場合はインストールします。 
2.  次のコマンドを実行します。XSD_FILE をマニフェスト XSD ファイルへのパスに置き換え、XML_FILE をマニフェスト XML ファイルへのパスに置き換えます。
    ```
    xmllint --noout --schema XSD_FILE XML_FILE
    ```

## <a name="use-runtime-logging-to-debug-your-add-in-manifest"></a>ランタイム ログを使用して、アドイン マニフェストをデバッグする

ランタイムのログを使用して、アドインのマニフェストをデバッグできます。この機能は、リソース ID の不一致のような XSD スキーマ検証では検出されないマニフェストの問題を識別して修正するのに役立ちます。ランタイムのログは、アドイン コマンドを実装するアドインのデバッグに特に有効です。  

>**注:**ランタイムのログ機能は現在、Office 2016 デスクトップで利用可能です。

### <a name="turn-on-runtime-logging"></a>ランタイムのログを有効にする

>**重要**: ランタイムのログはパフォーマンスに影響します。アドイン マニフェストに関する問題をデバッグする必要がある場合にのみ有効にしてください。

1. Office 2016 デスクトップのビルド **16.0.7019** 以降を実行していることを確認します。 
2. 'HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\' にレジストリ キー `RuntimeLogging` を追加します。 
3. キーの既定値にログを書き込むファイルの完全なパスを設定します。例については、[EnableRuntimeLogging.zip](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/raw/master/Tools/RuntimeLogging/EnableRuntimeLogging.zip) を参照してください。 

 > **注:**ログ ファイルが書き込まれるディレクトリが既に存在しており、書き込みアクセス許可がある必要があります。 
 
レジストリは次の図のようになります。![RuntimeLogging レジストリ キーを追加したレジストリ エディターのスクリーンショット](http://i.imgur.com/Sa9TyI6.png)

この機能を無効にするには、`RuntimeLogging` キーをレジストリから削除します。 

### <a name="troubleshoot-issues-with-your-manifest"></a>マニフェストの問題をトラブルシューティングする

ランタイムのログを使ってアドインの読み込みに関する問題のトラブルシューティングを行うには、次のようにします。
 
1. テスト用に[アドインをサイドロード](sideload-office-add-ins-for-testing.md)します。 

    >注:ログ ファイルのメッセージ数を減らすため、テストするアドインのみをサイドロードすることをお勧めします。
2. 何も起こらず、アドインが表示されない (アドイン ダイアログ ボックスにも表示されない) 場合は、ログ ファイルを開きます。
3. ログ ファイルでアドインの ID を検索します。ID はマニフェストで定義します。ログ ファイルでは、この ID には `SolutionId` というラベルが付いています。 

次の例のログ ファイルでは、存在しないリソース ファイルを参照しているコントロールが示されています。この例の問題を修正するには、マニフェストの入力ミスを訂正するか、足りないリソースを追加します。

![見つからないリソース ID を指定するエントリが含まれるログ ファイルのスクリーンショット](http://i.imgur.com/f8bouLA.png) 

### <a name="known-issues-with-runtime-logging"></a>ランタイムのログに関する既知の問題

混乱を招くメッセージまたは正しく分類されていないメッセージがログ ファイルに書き込まれることがあります。たとえば次のような場合です。

- メッセージ "`Medium   Current host not in add-in's host list`" に続く "`Unexpected Parsed manifest targeting different host`" は、誤ってエラーとして分類されています。
- SolutionId が含まれていないメッセージ "`Unexpected    Add-in is missing required manifest fields  DisplayName`" は、多くの場合、エラーはデバッグ対象のアドインと関係ありません。 
- `Monitorable` メッセージは、システムの観点からのエラーと予想されます。場合によっては、スキップされたがマニフェスト失敗の原因にはならなかったスペル ミスのある要素のような、マニフェストの問題を示していることがあります。 

## <a name="clear-the-office-cache"></a>Office のキャッシュをクリアする

リボン ボタンのアイコンのファイル名やアドイン コマンドのテキストなど、マニフェスト ファイルに変更を加えたときに、変更内容が反映されていないと思われる場合は、そのコンピューターで Office のキャッシュをクリアしてみてください。 

#### <a name="for-windows"></a>Windows の場合: 
フォルダー `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` の内容を削除します。

#### <a name="for-mac"></a>Mac の場合: 
フォルダー `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/` の内容を削除します。

#### <a name="for-ios"></a>iOS の場合: 
アドイン内の JavaScript から `window.location.reload(true)` を呼び出して強制的に再読み込みします。または、Office を再インストールしてください。

## <a name="additional-resources"></a>追加リソース

- [Office アドインの XML マニフェスト](../overview/add-in-manifests.md)
- [テスト用に Office アドインをサイドロードする](sideload-office-add-ins-for-testing.md)
- [Office アドインをデバッグする](debug-add-ins-using-f12-developer-tools-on-windows-10.md)