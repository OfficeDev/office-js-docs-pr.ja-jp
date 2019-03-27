---
title: iPad と Mac で Office アドインをデバッグする
description: ''
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: 5bf626c4c18bcedccd331570b6b892a8c6a903fd
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870402"
---
# <a name="debug-office-add-ins-on-ipad-and-mac"></a>iPad と Mac で Office アドインをデバッグする

Windows でのアドインの開発とデバッグには Visual Studio を使用できますが、iPad と Mac で使用して アドインをデバッグすることはできません。アドインは HTML と Javascript を使用して開発されているため、さまざまなプラットフォームで機能するように設計されていますが、さまざまなブラウザーで HTML の表示方法に微妙な違いがあります。この記事では、iPad または Mac で動作するアドインをデバッグする方法を説明します。

## <a name="debugging-with-vorlonjs-on-ipad-or-mac"></a>iPad または Mac での Vorlon.JS を使用したデバッグ

iPad または Mac でアドインをデバッグするには、Vorlon.JS (F12 ツールに似ている Web ページのデバッガー) を使用できます。 リモートで動作するように設計されているため、異なるデバイス間で Web ページをデバッグすることができます。 詳細については、[Vorlon の Web サイト](http://www.vorlonjs.com)を参照してください。  


### <a name="install-and-set-up-vorlonjs"></a>Vorlon をインストールしてセットアップする  

1.  管理者としてデバイスにログオンします。

2.  まだ [Node.js](https://nodejs.org) をインストールしていない場合は、インストールします。

3.  **[ターミナル]** ウィンドウを開き、コマンド `npm i -g vorlon` を入力します。ツールが `/usr/local/lib/node_modules/vorlon` にインストールされます。


### <a name="configure-vorlonjs-to-use-https"></a>Vorlon.JS を構成して HTTPS を使用する

Vorlon.JS を使用してアプリケーションをデバッグするには、既知の場所から Vorlon.JS スクリプトを読み込むアプリケーションの開始ページに `<script>` タグを追加します (詳細については、次の手順を参照してください)。アドインが SSL 保護付き (HTTPS) の場合、アドインで使用するすべてのスクリプトは HTTPS サーバーからホストされるように拡張する必要があります。これには、Vorlon.JS スクリプトも含まれます。そのため、アドインで Vorlon.JS を使用するには、Vorlon.JS を構成して SSL を使用することが必要になります。

> [!IMPORTANT]
> [!include[HTTPS guidance](../includes/https-guidance.md)]

1.  **[Finder]** で、`/usr/local/lib/node_modules/vorlon` に移動し、`/Server` フォルダーのコンテキスト メニュー (右クリック) を開き、**[情報を見る]** を選択します。

2.  **[サーバー情報]** ウィンドウの右下隅にある南京錠アイコンを選択して、フォルダーのロックを解除します。

3. ウィンドウの **[共有とアクセス権]** セクションで、**スタッフ** グループの **[特権]** を **[読み取り/書き込み]** に設定します。

4. 南京錠アイコンをもう一度選択して、フォルダーを***再度ロック***します。

5. **[Finder]** に戻り、`/Server` サブフォルダーを展開し、ファイル `config.json` を右クリックして、**[情報を見る]** を選択します。

6. **[config.json 情報]** ウィンドウで、親 `/Server` フォルダーに対して行ったものと同じ方法でファイルの特権を変更します。必ず再度ロックしてからウィンドウを閉じてください。

7. **[Finder]** に戻り、ファイル `config.json` を右クリックして、**[このアプリケーションで開く]**、**[テキストエディット]** の順に選択します。ファイルがテキスト エディターで開きます。

8. **useSSL** プロパティの値を `true` に変更します。

9. **[プラグイン]** セクションで、**ID** が `OFFICE` で**名前**が `Office Addin` のプライグインを検索します。プラグインの **enabled** プロパティがまだ `true` になっていない場合は、`true` に設定します。

10. ファイルを保存し、エディターを閉じます。

11. **[検索]** で `/usr/local/lib/node_modules/vorlon` に移動して、`Server` サブフォルダーを右クリックし、**[フォルダーの新しいターミナル]** を選択します。

12. **[ターミナル]** ウィンドウで、`sudo vorlon` と入力します。管理者パスワードの入力を求めるダイアログ ボックスが表示されます。Vorlon サーバーが起動します。**[ターミナル]** ウィンドウを開いたままにしておきます。

13. ブラウザー ウィンドウを開き、Vorlon.JS インターフェイスの `https://localhost:1337` に進みます。ダイアログ ボックスが表示されたら、**[常に]** を選択して、セキュリティ証明書を信頼します。

    > [!NOTE]
    > ダイアログ ボックスが表示されない場合は、手動で証明書を信頼する必要があります。証明書ファイルは `/usr/local/lib/node_modules/vorlon/Server/cert/server.crt` です。次の手順を実行し、問題が発生した場合は、Macintosh または iPad のヘルプを参照してください。
    >
    > 1. ブラウザー ウィンドウを閉じ、Vorlon サーバーを実行している **[ターミナル]** ウィンドウで、Control-C を使用してサーバーを停止します。
    > 2. **[Finder]** で、`server.crt` ファイルを右クリックして、**[キーチェーンアクセス]** を選択します。**[キーチェーンアクセス]** ウィンドウが開きます。
    > 3. 左側の **[キーチェーン]** リストで、**[ログイン]** がまだ選択されていない場合は選択し、**[カテゴリ]** セクションで **[証明書]** を選択します。証明書 **localhost** が一覧表示されます。
    > 4. 証明書 **localhost** を右クリックし、**[情報を見る]** を選択します。**[localhost]** ウィンドウが開きます。
    > 5. **[信頼]** セクションで、**[この証明書を使用する場合]** というラベルの付いたセレクターを選択し、**[常に信頼する]** を選択します。 
    > 6. **[localhost]** ウィンドウを閉じます。アクションが成功すると、**[キーチェーンアクセス]** ウィンドウの **localhost** 証明書のアイコンに青い円で囲まれた白い十字が表示されます。


### <a name="configure-the-add-in-for-vorlonjs-debugging"></a>Vorlon.JS デバッグ用のアドインを構成します。

1. 次のスクリプト タグを、アドインの home.html ファイル (またはメイン HTML ファイル) の `<head>` セクションに追加します。

    ```html
    <script src="https://localhost:1337/vorlon.js"></script>
    ```  

2. Azure Web サイトなど、Mac または iPad からアクセス可能な Web サーバーにアドイン Web アプリケーションを展開します。

3. アドイン マニフェストに URL が表示されるすべての場所で、アドインの URL を更新します。

4. アドイン マニフェストを Mac または iPad 上の次のフォルダーにコピーします: `/Users/{your_name_on_the_device}/Library/Containers/com.microsoft.{host_name}/Data/Documents/wef`。ここで、*{host_name}* は、Word、Excel、PowerPoint、または Outlook です。


### <a name="inspect-an-add-in-in-vorlonjs"></a>Vorlon.JS でアドインを検査する

1. Vorlon サーバーが実行されていない場合、**[Finder]** で `/usr/local/lib/node_modules/vorlon` に移動して、`Server` サブフォルダーを右クリックし、**[フォルダーの新しいターミナル]** を選択します。 

2.  **[ターミナル]** ウィンドウで、`sudo vorlon` と入力します。管理者パスワードの入力を求めるダイアログ ボックスが表示されます。Vorlon サーバーが起動します。**[ターミナル]** ウィンドウを開いたままにしておきます。

3.  ブラウザー ウィンドウを開き、Vorlon.JS インターフェイスの `https://localhost:1337` に進みます。

4. アドインをサイドロードします。 アドインが Excel、PowerPoint、Word 用の場合は、「[iPad または Mac で Office アドインをサイドロードする](sideload-an-office-add-in-on-ipad-and-mac.md)」の説明に従ってサイドロードします。 アドインが Outlook アドインである場合は、「[テストのために Outlook アドインをサイドロードする](/outlook/add-ins/sideload-outlook-add-ins-for-testing)」の説明に従ってサイドロードします。 アドインでアドイン コマンドを使用しない場合は、アドインが直ちに開きます。 それ以外の場合は、ボタンを選択してアドインを開きます。 Office ホスト アプリケーションのビルドに応じて、ボタンは **[ホーム]** タブまたは **[アドイン]** タブのいずれかに表示されます。

アドインは、Vorlon.JS のクライアントのリスト (Vorlon.JS インターフェイスの左側) に **{OS} - n** として表示されます。*n* は数値、*{OS}* は "Macintosh" などのデバイスの種類です。

![Vorlon.js インターフェイスを示すスクリーンショット](../images/vorlon-interface.png)

Vorlon ツールには、さまざまなプラグインがあります。現在有効になっているプラグインはツールの上部にタブとして表示されます。 (左側にある歯車アイコンを選択すると、さらに別のプラグインを有効にすることができます。)これらのプラグインは、F12 ツールの機能に似ています。 たとえば、DOM 要素の強調表示、コマンドの実行などを行えます。 詳細については、[Vorlon ドキュメントの「コア プラグイン」](http://vorlonjs.com/documentation/#console)を参照してください。

**Office アドイン** プラグインにより Office.js に特別な機能 (オブジェクト モデルを調査する機能、Office.js の呼び出しを実行する機能、およびオブジェクト プロパティの値を読み取る機能など) が追加されます。手順については、「[Office アドインをデバッグするための VorlonJS プラグイン](https://blogs.msdn.microsoft.com/mim/2016/02/18/vorlonjs-plugin-for-debugging-office-addin/)」を参照してください。

> [!NOTE]
> Vorlon.JS にブレーク ポイントを設定する方法はありません。

## <a name="debugging-with-safari-web-inspector-on-a-mac"></a>Mac での Safari Web インスペクタを使用したデバッグ

> [!IMPORTANT]
> **要素の検査**アドイン コンテキスト メニュー オプションは試験的な機能であり、Office アプリケーションの将来のバージョンでこの機能が維持されるかどうかは保証されない点に注意してください。

作業ウィンドウまたはコンテンツ アドインに UI を表示するアドインを使用している場合は、Safari Web インスペクタを使用して Office アドインをデバッグできます。

Mac の Office アドインをデバッグするには、Mac OS High Sierra 以降 と Mac Office バージョン 16.9.1 (ビルド 18012504) 以降の両方が必要です。 Office for Mac ビルドをまだお持ちでない場合は、[Office 365 Developer Program](https://aka.ms/o365devprogram) に参加することで入手できます。

最初に端末を開き、該当する Office アプリケーションの `OfficeWebAddinDeveloperExtras` プロパティを以下のように設定します。

- `defaults write com.microsoft.Word OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Excel OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Powerpoint OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Outlook OfficeWebAddinDeveloperExtras -bool true`

次に Office アプリケーションを開き、[アドインをサイドロードします](sideload-an-office-add-in-on-ipad-and-mac.md)。 アドインを右クリックします。コンテキスト メニューに **[要素の検査]** オプションが表示されるはずです。  このオプションを選択するとインスペクタが表示されます。インスペクタでは、ブレークポイントを設定してアドインをデバッグできます。

> [!NOTE]
> インスペクタを使用するとダイアログのちらつきが発生する場合は、次の回避策を試してください。
> 1. ダイアログのサイズを変更します。
> 2. **[要素の検査]** を選択します (新しいウィンドウが開きます)。
> 3. ダイアログを元のサイズに変更します。
> 4. 必要に応じてインスペクタを使用します。


## <a name="clearing-the-office-applications-cache-on-a-mac-or-ipad"></a>Mac または iPad 上の Office アプリケーションのキャッシュのクリア

アドインはパフォーマンス上の理由から、Office for Mac でキャッシュされることが多いです。通常、キャッシュはアドインを再読み込みすることでクリアされます。同じドキュメント内に複数のアドインが存在する場合、再読み込み時にキャッシュを自動的にクリアするプロセスは信頼できない場合があります。

Mac では、`/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/` フォルダー内にあるすべてを削除することによってキャッシュを手動でクリアできます。

iPad では、アドインの JavaScript から `window.location.reload(true)` を呼び出して、強制的に再読み込みすることができます。または、Office を再インストールすることができます。
