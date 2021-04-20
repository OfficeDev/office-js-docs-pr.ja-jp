---
title: Visual Studio Code と Azure を使用してアドインを発行する
description: Visual Studio Code と Azure Active Directory を使用してアドインを発行する方法
ms.date: 08/12/2020
localization_priority: Normal
ms.openlocfilehash: 3552e4eebacc84fc2b8e37782c97b4e03e96e508
ms.sourcegitcommit: 7faa0932b953a4983a80af70f49d116c3236d81a
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/21/2020
ms.locfileid: "46845513"
---
# <a name="publish-an-add-in-developed-with-visual-studio-code"></a>Visual Studio Code で開発されたアドインを発行する

この記事では、Yeoman ジェネレーターを使用して作成し、[Visual Studio Code (VS Code)](https://code.visualstudio.com) またはその他のエディターで開発した Office アドインを発行する方法について説明します。

> [!NOTE]
> Visual Studio を使用して作成した Office アドインの発行の詳細については、「[Visual Studio を使用してアドインを発行する](package-your-add-in-using-visual-studio.md)」を参照してください。

## <a name="publishing-an-add-in-for-other-users-to-access"></a>他のユーザーがアクセスできるようにアドインを発行する

Office アドインは、Web アプリケーションとマニフェスト ファイルで構成されています。 Web アプリケーションはアドインのユーザー インターフェイスと機能を定義し、マニフェストは Web アプリケーションの場所を指定し、アドインの設定と機能を定義します。

を開発しているときに、ローカル web サーバーでアドインを実行できます ( `localhost` )。 他のユーザーがアクセスできるように公開する準備ができたら、web アプリケーションを展開し、マニフェストを更新して、展開されたアプリケーションの URL を指定する必要があります。

アドインが目的どおりに動作している場合は、Azure ストレージ拡張機能を使用して、Visual Studio Code を使用して直接発行できます。

## <a name="using-visual-studio-code-to-publish"></a>Visual Studio Code を使用して発行する

>[!NOTE]
> これらの手順は、[ごみ箱のジェネレーターを使用して作成されたプロジェクトに対してのみ機能します。

1. Visual Studio Code (VS コード) で、ルートフォルダーからプロジェクトを開きます。
2. VS Code の [Extensions] ビューで、Azure ストレージ拡張機能を検索してインストールします。
3. インストールされると、Azure アイコンがアクティビティバーに追加されます。 この拡張機能にアクセスするには、このチェックボックスをオンにします。 アクティビティバーが非表示の場合、拡張機能にアクセスすることはできません。 [> の表示 > 表示 **]** を選択してアクティビティバーを表示します。
4. 拡張機能を使用している場合は、[ **azure にサインイン**] を選択して azure アカウントにサインインします。 Azure アカウントをまだ持っていない場合は、[azure アカウント **を作成**する] を選択して、azure アカウントを作成することもできます。 提供される手順に従って、アカウントをセットアップします。
5. Azure アカウントにサインインすると、拡張機能に Azure storage アカウントが表示されます。 ストレージアカウントをまだ持っていない場合は、[ **新しいストレージアカウントの作成** ] オプションを使用して作成する必要があります。 ストレージアカウントに、「a-z」と「0-9」のみを使用して、グローバルに一意の名前を指定します。 既定では、これによってストレージアカウントとリソースグループが同じ名前で作成されることに注意してください。 これにより、自動的にストレージアカウントが West 米に配置されます。 これは [、Azure アカウント](https://portal.azure.com/)を使用してオンラインで調整できます。
6. ストレージアカウントを選択して保持 (右クリック) し、[ **静的 web サイトの構成**] を選択します。 インデックスドキュメント名と404ドキュメント名を入力するように求められます。 Index ドキュメント名を既定のからから `index.html` に変更し **`taskpane.html`** ます。 404のドキュメント名も変更する必要がありますが、にする必要はありません。
7. ストレージを選択して保持 (右クリック) し、今度は [ **静的 web サイトの参照**] を選択します。 開いたブラウザーウィンドウで、web サイトの URL をコピーします。
8. VS Code で、プロジェクトのマニフェストファイル () を開き、 `manifest.xml` localhost の url (など) への参照を `https://localhost:3000` コピーした url に変更します。 このエンドポイントは、新しく作成されたストレージアカウントの静的な web サイトの URL です。 マニフェストファイルへの変更を保存します。
9. コマンドラインプロンプトを開き、アドインプロジェクトのルートディレクトリに移動します。 その後、次のコマンドを実行して、運用展開用のすべてのファイルを準備します。

    ```command&nbsp;line
    npm run build
    ```

    ビルドが完了すると、アドイン プロジェクトのルート ディレクトリにある **dist** フォルダーに、以降の手順で展開するファイルが含まれます。

10. を展開するには、[ファイルエクスプローラー] を選択し、[ **dist** ] フォルダーを右クリックして、[ **静的 Web サイトに展開**] を選択します。 メッセージが表示されたら、前に作成したストレージアカウントを選択します。

![静的 web サイトへの展開](../images/deploy-to-static-website.png)

11. 展開が完了すると、 **web サイトを参照** するメッセージが表示され、展開されたアプリコードのプライマリエンドポイントを開くことができます。

## <a name="see-also"></a>関連項目

- [Visual Studio Code を使用して Office アドインを開発する](../develop/develop-add-ins-vscode.md)
- [Office アドインを展開し、発行する](../publish/publish.md)
