---
title: Visual Studio Code と Azure を使用してアドインを発行する
description: Visual Studio Code と Azure Active Directory を使用してアドインを発行する方法
ms.date: 09/07/2022
ms.custom: vscode-azure-extension-update-completed
ms.localizationpriority: medium
ms.openlocfilehash: f1e931042f7796333bbc75b2666f3564e63b9f7b
ms.sourcegitcommit: 005783ddd43cf6582233be1be6e3463d7ab9b0e5
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/05/2022
ms.locfileid: "68467197"
---
# <a name="publish-an-add-in-developed-with-visual-studio-code"></a>Visual Studio Code で開発されたアドインを発行する

この記事では、Yeoman ジェネレーターを使用して作成し、[Visual Studio Code (VS Code)](https://code.visualstudio.com) またはその他のエディターで開発した Office アドインを発行する方法について説明します。

> [!NOTE]
> Visual Studio を使用して作成した Office アドインの発行の詳細については、「[Visual Studio を使用してアドインを発行する](package-your-add-in-using-visual-studio.md)」を参照してください。

## <a name="publishing-an-add-in-for-other-users-to-access"></a>他のユーザーがアクセスできるようにアドインを発行する

An Office Add-in consists of a web application and a manifest file. The web application defines the add-in's user interface and functionality, while the manifest specifies the location of the web application and defines settings and capabilities of the add-in.

開発中は、ローカル Web サーバー (`localhost`) でアドインを実行できます。 他のユーザーがアクセスできるように公開する準備ができたら、Web アプリケーションをデプロイし、マニフェストを更新して、デプロイされたアプリケーションの URL を指定する必要があります。

アドインが必要に応じて動作している場合は、Azure Storage 拡張機能を使用して Visual Studio Code から直接発行できます。

## <a name="using-visual-studio-code-to-publish"></a>Visual Studio コードを使用して発行する

>[!NOTE]
> これらの手順は、Yeoman ジェネレーターで作成されたプロジェクトでのみ機能します。

1. Visual Studio Code (VS Code) のルート フォルダーからプロジェクトを開きます。
1. [ **拡張機能の表示] >** 選択して (Ctrl + Shift + X) 拡張機能ビューを開きます。
1. **Azure Storage** 拡張機能を検索してインストールします。
1. インストールが完了すると、 **アクティビティ バー** に Azure アイコンが追加されます。 拡張機能にアクセスするには、それを選択します。 **アクティビティ バー** が非表示の場合は、[**表示>外観>アクティビティ バー**] を選択して開きます。
1. [ **Azure へのサインイン** ] を選択して、Azure アカウントにサインインします。 Azure アカウントをまだ持っていない場合は、[Azure アカウントの作成] を選択して作成 **します**。 指定した手順に従ってアカウントを設定します。

    :::image type="content" source="../images/azure-extension-sign-in.png" alt-text="Azure 拡張機能で選択されている [Azure にサインイン] ボタンをクリックします。":::

1. サインインすると、拡張機能に Azure ストレージ アカウントが表示されます。 ストレージ アカウントをまだ持っていない場合は、コマンド パレットの **[ストレージ アカウントの作成** ] オプションを使用して作成します。 ストレージ アカウントには、"a-z" と "0- 9" のみを使用してグローバルに一意の名前を付けます。 既定では、ストレージ アカウントとリソース グループが同じ名前で作成されることに注意してください。 ストレージ アカウントは米国西部に自動的に格納されます。 これは、 [Azure アカウント](https://portal.azure.com/)を介してオンラインで調整できます。

    :::image type="content" source="../images/azure-extension-create-storage-account.png" alt-text="Azure 拡張機能でストレージ アカウントを作成>ストレージ アカウントを選択します。":::

1. ストレージ アカウントを右クリックし、[ **静的 Web サイトの構成**] を選択します。 インデックス ドキュメント名と 404 ドキュメント名を入力するように求められます。 インデックス ドキュメント名を既定値`index.html`**`taskpane.html`** から . 404 ドキュメント名を変更することもできますが、変更する必要はありません。
1. ストレージ アカウントをもう一度右クリックし、今度は **[静的 Web サイトの参照**] を選択します。 開いたブラウザー ウィンドウから、Web サイトの URL をコピーします。
1. プロジェクトのマニフェスト ファイル (`manifest.xml`) を開き、localhost URL (など `https://localhost:3000`) への参照をすべてコピーした URL に変更します。 このエンドポイントは、新しく作成したストレージ アカウントの静的 Web サイト URL です。 変更をマニフェスト ファイルに保存します。
1. コマンド ライン プロンプトまたはターミナル ウィンドウを開き、アドイン プロジェクトのルート ディレクトリに移動します。 次のコマンドを実行して、運用環境のデプロイ用にすべてのファイルを準備します。

    ```command&nbsp;line
    npm run build
    ```

    ビルドが完了すると、アドイン プロジェクトのルート ディレクトリにある **dist** フォルダーに、以降の手順で展開するファイルが含まれます。

1. VS Code でエクスプローラーに移動し、 **dist** フォルダーを右クリックし、 **Azure Storage 経由で [静的 Web サイトにデプロイ]** を選択します。 メッセージが表示されたら、前に作成したストレージ アカウントを選択します。

    :::image type="content" source="../images/deploy-to-static-website.png" alt-text="dist フォルダーを選択し、右クリックして[Azure Storage 経由で静的 Web サイトにデプロイ] を選択します。":::

1. デプロイが完了したら、前に作成したストレージ アカウントを右クリックし、[ **静的 Web サイトの参照**] を選択します。 静的 Web サイトが開き、作業ウィンドウが表示されます。

1. 最後に、 [マニフェスト ファイルをサイドロード](../testing/sideload-office-add-ins-for-testing.md) すると、先ほどデプロイした静的 Web サイトからアドインが読み込まれます。

## <a name="deploy-custom-functions-for-excel"></a>Excel 用のカスタム関数をデプロイする

アドインにカスタム関数がある場合は、Azure Storage アカウントで有効にする手順がいくつかあります。 まず、OFFICE が functions.json ファイルにアクセスできるように CORS を有効にします。

1. Azure ストレージ アカウントを右クリックし、[ **ポータルで開く**] を選択します。
1. [設定] グループで、[ **リソース共有 (CORS)]** を選択します。 検索ボックスを使用して、これを見つけることもできます。
1. 次の設定で新しい CORS ルールを作成します。

    |プロパティ        |値                        |
    |----------------|-----------------------------|
    |許可される配信元 | \*                          |
    |許可されるメソッド | GET                         |
    |許可されるヘッダー | \*                          |
    |公開されたヘッダー | Access-Control-Allow-Origin |
    |最大有効期間         | 200                         |

1. **[保存]** を選択します。

> [!CAUTION]
> この CORS 構成では、サーバー上のすべてのファイルがすべてのドメインで一般公開されていることを前提としています。  

次に、JSON ファイルの MIME の種類を追加します。

1. web.configという名前の /src フォルダーに新しいファイル **を** 作成します。
1. 次の XML を挿入し、ファイルを保存します。

    ```xml
    <?xml version="1.0"?>
    <configuration>
      <system.webServer>
        <staticContent>
          <mimeMap fileExtension=".json" mimeType="application/json" />
        </staticContent>
      </system.webServer>
    </configuration> 
    ```

1. **webpack.config.js** ファイルを開きます。
1. ビルドの実行時にバンドルにweb.configをコピーするには、次の `plugins` コードを一覧に追加します。

    ```javascript
    new CopyWebpackPlugin({
      patterns: [
      {
        from: "src/web.config",
        to: "src/web.config",
      },
     ],
    }),
    ```

1. コマンド ライン プロンプトを開き、アドイン プロジェクトのルート ディレクトリに移動します。 次に、次のコマンドを実行して、すべてのファイルをデプロイ用に準備します。

    ```command&nbsp;line
    npm run build
    ```

    ビルドが完了すると、アドイン プロジェクトのルート ディレクトリ内の **dist** フォルダーに、展開するファイルが含まれます。

1. デプロイするには、VS Code **Explorer** で **dist** フォルダーを右クリックし、 **Azure Storage 経由で [静的 Web サイトにデプロイ]** を選択します。 メッセージが表示されたら、前に作成したストレージ アカウントを選択します。 **dist** フォルダーを既にデプロイしている場合は、Azure Storage 内のファイルを最新の変更で上書きするかどうかを確認するメッセージが表示されます。

## <a name="see-also"></a>関連項目

- [Visual Studio Code を使用して Office アドインを開発する](../develop/develop-add-ins-vscode.md)
- [Office アドインを展開し、発行する](../publish/publish.md)
- [Azure Storage のクロスオリジン リソース共有 (CORS) のサポート](/rest/api/storageservices/cross-origin-resource-sharing--cors--support-for-the-azure-storage-services)
