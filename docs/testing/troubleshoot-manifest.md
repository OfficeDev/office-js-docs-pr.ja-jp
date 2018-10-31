---
title: マニフェストの問題を検証し、トラブルシューティングを行う
description: 以下の方法を使用して、Office アドイン マニフェストを検証します。
ms.date: 12/04/2017
ms.openlocfilehash: 51d644f7cfb7fbad5c9b66be41dc57015202b9be
ms.sourcegitcommit: a6d6348075c1abed76d2146ddfc099b0151fe403
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/19/2018
ms.locfileid: "25639988"
---
# <a name="validate-and-troubleshoot-issues-with-your-manifest"></a>マニフェストの問題を検証し、トラブルシューティングを行う

以下の方法を使用して、Office アドイン マニフェストの問題を検証し、トラブルシューティングを行います。 

- [Office アドイン検証ツールを使用してマニフェストを検証する](#validate-your-manifest-with-the-office-add-in-validator)   
- [XML スキーマと比較してマニフェストを検証する](#validate-your-manifest-against-the-xml-schema)
- [ランタイム ログを使用して、アドイン マニフェストをデバッグする](#use-runtime-logging-to-debug-your-add-in-manifest)


## <a name="validate-your-manifest-with-the-office-add-in-validator"></a>Office アドイン検証ツールを使用してマニフェストを検証する

Office アドインを記述するマニフェスト ファイルが正確かつ完全であることを確認するために、[Office アドイン検証ツール](https://github.com/OfficeDev/office-addin-validator)を使用してマニフェスト ファイルを検証します。

### <a name="to-use-the-office-add-in-validator-to-validate-your-manifest"></a>Office アドイン検証ツールを使用してマニフェストを検証するには

1. [Node.js](https://nodejs.org/download/) をインストールします。 

2. 管理者としてコマンド プロンプト/ターミナルを開き、次のコマンドを使用して Office アドイン検証ツールとその依存関係をグローバルにインストールします。

    ```bash
    npm install -g office-addin-validator
    ```
    
    > [!NOTE]
    > Yo Office が既にインストールされている場合、最新のバージョンにアップグレードすると、検証ツールが依存関係としてインストールされます。

3. マニフェストを検証するには、次のコマンドを実行し、MANIFEST.XML をマニフェスト XML ファイルへのパスに置き換えます。

    ```bash
    validate-office-addin MANIFEST.XML
    ```

## <a name="validate-your-manifest-against-the-xml-schema"></a>XML スキーマと比較してマニフェストを検証する

マニフェスト ファイルが、使用している要素の名前空間を含む正しいスキーマに従っていることを確認します。サンプル マニフェストから要素をコピーした場合は、**適切な名前空間が含まれる**か再度確認してください。マニフェストは [XML スキーマ定義 (XSD)](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas) ファイルに照らして検証することができます。XML スキーマの検証ツールを使用すると、この検証を実行します。 



### <a name="to-use-a-command-line-xml-schema-validation-tool-to-validate-your-manifest"></a>コマンドライン XML スキーマ検証ツールを使用してマニフェストを検証するには

1.  [tar](https://www.gnu.org/software/tar/) および [libxml](http://xmlsoft.org/FAQ.html) をまだインストールしていない場合はインストールします。

2.  次のコマンドを実行します。`XSD_FILE` をマニフェスト XSD ファイルへのパスに置き換え、`XML_FILE` をマニフェスト XML ファイルへのパスに置き換えます。
    
    ```bash
    xmllint --noout --schema XSD_FILE XML_FILE
    ```

## <a name="use-runtime-logging-to-debug-your-add-in"></a>ランタイムのログを使用して、アドインをデバッグする 

アドインのマニフェストといくつかのインストール エラーをデバッグするには、ランタイムのログが使用できます。この機能を使用すると、リソース ID の不一致など、XSD スキーマ検証では検出されない問題を修正することができます。ランタイムのログは、アドインのコマンドと Excel のカスタム関数を実装するアドインをデバッグするために特に便利です。   

> [!NOTE]
> ランタイムのログ機能は現在、Office 2016 デスクトップで利用可能です。

### <a name="to-turn-on-runtime-logging"></a>ランタイムのログを有効にするには

> [!IMPORTANT]
> ランタイムのログはパフォーマンスに影響します。アドイン マニフェストに関する問題をデバッグする必要がある場合にのみ有効にしてください。

ランタイムのログを有効にするには、以下を実行します。

1. Office 2016 デスクトップのビルド **16.0.7019** 以降を実行していることを確認します。 

2. `RuntimeLogging` の下にレジストリ キーを追加します`HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\` 。 

3. キーの既定値にログを書き込むファイルの完全なパスを設定します。例については、[EnableRuntimeLogging.zip](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/raw/master/Tools/RuntimeLogging/EnableRuntimeLogging.zip) を参照してください。 

    > [!NOTE]
    > ログ ファイルが書き込まれるディレクトリが既に存在し、書き込みアクセス許可がある必要があります。 
 
レジストリは次の図のようになります。この機能を無効にするには、`RuntimeLogging`キーをレジストリから削除します。 

![RuntimeLogging レジストリ キーを追加したレジストリ エディターのスクリーンショット](http://i.imgur.com/Sa9TyI6.png)


### <a name="to-troubleshoot-issues-with-your-manifest"></a>マニフェストの問題のトラブルシューティングを行うには

ランタイムのログを使用してアドインの読み込みに関する問題のトラブルシューティングを行うには、次のようにします。
 
1. テスト用に[アドインをサイドロード](sideload-office-add-ins-for-testing.md)します。 

    > [!NOTE]
    > ログ ファイルのメッセージ数を最小限に抑えるため、テストするアドインのみをサイドロードすることをお勧めします。

2. 何も起こらず、アドインが表示されない (アドイン ダイアログ ボックスにも表示されない) 場合は、ログ ファイルを開きます。

3. ログ ファイルでアドインの ID を検索します。ID はマニフェストで定義します。ログ ファイル内で、この ID には `SolutionId` というラベルが付いています。 

次の例のログ ファイルでは、存在しないリソース ファイルを参照しているコントロールが示されています。この例の問題を修正するには、マニフェストの入力ミスを訂正するか、足りないリソースを追加します。

![見つからないリソース ID を指定するエントリが含まれるログ ファイルのスクリーンショット](http://i.imgur.com/f8bouLA.png) 

### <a name="known-issues-with-runtime-logging"></a>ランタイムのログに関する既知の問題

混乱を招くメッセージまたは正しく分類されていないメッセージがログ ファイルに書き込まれることがあります。たとえば次のような場合です。

- メッセージ "`Medium Current host not in add-in's host list`" に続く "`Unexpected Parsed manifest targeting different host`" は、誤ってエラーとして分類されています。

- SolutionId が含まれていないメッセージ "`Unexpected Add-in is missing required manifest fields DisplayName`" は、多くの場合、エラーはデバッグ対象のアドインと関係ありません。 

- すべての `Monitorable`  メッセージは、システムの観点からのエラーと予想されます。場合によっては、スキップされたがマニフェスト失敗の原因にはならなかったスペル ミスのある要素をはじめとする、マニフェストの問題を示していることがあります。 

## <a name="clear-the-office-cache"></a>Office のキャッシュをクリアする

リボン ボタンのアイコンのファイル名やアドイン コマンドのテキストなど、マニフェスト ファイルに変更を加えたときに、変更内容が反映されていないと思われる場合は、そのコンピューターで Office のキャッシュをクリアしてみてください。 

#### <a name="for-windows"></a>Windows の場合:
フォルダー `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` の内容を削除します。

#### <a name="for-mac"></a>Mac の場合:
フォルダー `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/` の内容を削除します。

#### <a name="for-ios"></a>iOS の場合:
アドイン内の JavaScript から `window.location.reload(true)` を呼び出して強制的に再読み込みします。または、Office を再インストールしてください。

## <a name="see-also"></a>関連項目

- [Office アドインの XML マニフェスト](../develop/add-in-manifests.md)
- [テスト用に Office アドインをサイドロードする](sideload-office-add-ins-for-testing.md)
- [Office アドインをデバッグする](debug-add-ins-using-f12-developer-tools-on-windows-10.md)
