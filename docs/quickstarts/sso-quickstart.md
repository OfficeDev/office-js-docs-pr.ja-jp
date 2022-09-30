---
title: シングル サインオン (SSO) のクイック スタート
description: Yeoman ジェネレーターを使用して、シングル サインオンを使用する Node.js Office アドインを作成する。
ms.date: 09/07/2022
ms.prod: non-product-specific
ms.localizationpriority: high
ms.openlocfilehash: ecbecfd7e475c224451735c7a864f6de2c230d07
ms.sourcegitcommit: cff5d3450f0c02814c1436f94cd1fc1537094051
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/30/2022
ms.locfileid: "68239375"
---
# <a name="single-sign-on-sso-quick-start"></a>シングル サインオン (SSO) のクイック スタート

この記事では、Office アドイン用 Yeoman ジェネレーターを使用して、シングル サインオン (SSO) を使用する Excel、Outlook、Word、または PowerPoint 用の Office アドインを作成します。

> [!NOTE]
> Office アドイン用 Yeoman ジェネレーターによって提供される SSO テンプレートは、localhost でのみ実行され、展開できません。 運用目的で SSO を使用して新しい Office アドインをビルドする場合は、「 [シングル サインオンを使用する office アドインNode.js作成する](../develop/create-sso-office-add-ins-nodejs.md)」の手順に従います。

## <a name="prerequisites"></a>前提条件

- [Node.js](https://nodejs.org) (最新 [LTS](https://nodejs.org/about/releases) バージョン)。

- 最新バージョンの [Yeoman](https://github.com/yeoman/yo) と [Office アドイン用の Yeoman ジェネレーター](../develop/yeoman-generator-overview.md)。これらのツールをグローバルにインストールするには、コマンド プロンプトから次のコマンドを実行します。

    ```command&nbsp;line
    npm install -g yo generator-office
    ```

    [!include[note to update Yeoman generator](../includes/note-yeoman-generator-update.md)]

- Mac を使用していて Azure CLI がコンピューターにインストールされていない場合は、[Homebrew](https://brew.sh/) をインストールする必要があります。 このクイック スタート中に実行する SSO 構成スクリプトは、Homebrew を使用して Azure CLI をインストールした後、Azure CLI を使用して Azure の SSO を構成します。

## <a name="create-the-add-in-project"></a>アドイン プロジェクトの作成

> [!TIP]
> Yeoman ジェネレーターは、スクリプトの種類が JavaScript または TypeScript の Excel、Outlook、Word、または PowerPoint 用の SSO 対応 Office アドインを作成できます。 次の手順では、`JavaScript` と `Excel` を指定しますが、使用しているシナリオに最適なスクリプト タイプと Office クライアント アプリケーションを選択する必要があります。

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- **Choose a project type: (プロジェクトの種類を選択)** `Office Add-in Task Pane project supporting single sign-on (localhost)`
- **Choose a script type: (スクリプトの種類を選択)** `JavaScript`
- **What would you want to name your add-in?: (アドインの名前を何にしますか)** `My Office Add-in`
- **Which Office client application would you like to support?: (どの Office クライアント アプリケーションをサポートしますか)** 、、`Outlook``Word`またはを選択`Excel`します`Powerpoint`。

:::image type="content" source="../images/yo-office-sso-excel.png" alt-text="コマンド ライン インターフェイスで Yeoman ジェネレーターのプロンプトと回答を表示します。":::

ウィザードを完了すると、ジェネレーターによってプロジェクトが作成されて、サポートしているノード コンポーネントがインストールされます。

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="explore-the-project"></a>プロジェクトを確認する

Yeoman ジェネレーターで作成したアドイン プロジェクトには、SSO が有効な作業ウィンドウ アドインのコードが含まれています。

[!include[project structure for an SSO-enabled add-in created with the Yeoman generator](../includes/sso-yeoman-project-structure.md)]

## <a name="configure-sso"></a>SSO を構成する

アドイン プロジェクトが作成され、SSO プロセスを容易にするために必要なコードが含まれるので、次の手順に従ってアドインの SSO を構成します。

1. プロジェクトのルート フォルダーに移動します。

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

2. 次のコマンドを実行して、アドインの SSO を構成します。

    ```command&nbsp;line
    npm run configure-sso
    ```

    > [!WARNING]
    > このコマンドは、テナントが 2 要素認証を要求するように構成されている場合、失敗します。 このシナリオでは、 [シングル サインオンを使用する Office アドインの Node.js作成](../develop/create-sso-office-add-ins-nodejs.md) に関するチュートリアルの手順に従って、Azure アプリの登録と SSO 構成の手順を手動で完了する必要があります。

3. Web ブラウザー ウィンドウが開き、Azure にサインインするように指示されます。 Microsoft 365 管理者の資格情報を使用して Azure にサインインします。 これらの資格情報を使用して Azure に新しいアプリケーションが登録され、SSO に必要な設定が構成されます。

    > [!NOTE]
    > この手順で、管理者以外の資格情報を使用して Azure にサインインした場合、`configure-sso` スクリプトは、組織内のユーザーにアドインの管理者の同意を提供しません。 そのため、SSO はアドインのユーザーは使用できず、サインインするように求められます。

4. 資格情報を入力したら、ブラウザー ウィンドウを閉じ、コマンド プロンプトに戻ります。 SSO の構成プロセスが続行されると、コンソールに書き込まれたステータス メッセージが表示されます。 コンソール メッセージで説明されているように、Yeoman ジェネレーターが作成したアドイン プロジェクト内のファイルは、SSO プロセスで必要なデータで自動的に更新されます。

## <a name="test-your-add-in"></a>アドインをテストする

Excel、Word、または PowerPoint アドインを作成した場合は、次のセクションの手順に従って試してください。 Outlook アドインを作成した場合は、代わりに [Outlook](#outlook) セクションの手順を完了します。

### <a name="excel-word-and-powerpoint"></a>Excel、Word、および PowerPoint

Excel、Word、または PowerPoint アドインをテストするには、次の手順を実行します。

1. SSO の構成プロセスが完了したら、次のコマンドを実行してプロジェクトを構築し、ローカル Web サーバーを起動して以前に選択した Office クライアント アプリケーションにアドインをサイドロードします。

    [!INCLUDE [alert use https](../includes/alert-use-https.md)]

    ```command&nbsp;line
    npm start
    ```

2. 前のコマンドを実行したときに Excel、Word、または PowerPoint が開いたら、 [前のセクション](#configure-sso)の手順 3 で SSO の構成中に Azure への接続に使用した Microsoft 365 管理者アカウントと同じ Microsoft 365 組織のメンバーであるユーザー アカウントでサインインしていることを確認します。 これにより、SSO を正常に実行するための適切な条件が確立されます。

3. Office クライアント アプリケーションで 、[ **ホーム** ] タブを選択し、[ **タスクウィンドウの表示** ] を選択してアドイン作業ウィンドウを開きます。

    :::image type="content" source="../images/excel-quickstart-addin-3b.png" alt-text="Excel アドイン ボタン。":::

4. 作業ウィンドウの下部にある [**マイ ユーザー プロファイルの情報を取得する**] ボタンを選択して、SSO プロセスを開始します。

5. アドインの代わりにアクセス許可を要求するダイアログ ウィンドウが表示される場合は、SSO はシナリオでサポートされず、代わりにアドインが別のユーザー認証方法に戻っていることを意味します。 これは、テナント管理者が Microsoft Graph にアクセスするためのアドインの同意を付与していない場合、またはユーザーが有効な Microsoft アカウントまたはMicrosoft 365 Educationまたは職場アカウントを使用して Office にサインインしていない場合に発生する可能性があります。 [ **承諾]** を選択して続行します。

    ![[承認] ボタンが強調表示された [アクセス許可] 要求ダイアログを示すスクリーンショット。](../images/sso-permissions-request.png)

    > [!NOTE]
    > ユーザーがこのアクセス許可の要求を受け入れると、今後再びプロンプトが表示されることはありません。

6. アドインは、サインインしたユーザーのプロファイル情報を取得し、ドキュメントに書き込みます。 次の画像は、Excel ワークシートに書き込まれたプロファイル情報の例を示します。

    ![Excel ワークシートのユーザー プロファイル情報を示すスクリーンショット。](../images/sso-user-profile-info-excel.png)

### <a name="outlook"></a>Outlook

Outlook アドインを試すには、次の手順を実行します。

1. SSO 構成プロセスが完了したら、次のコマンドを実行してプロジェクトを構築し、ローカル Web サーバーを起動します。

    [!INCLUDE [alert use https](../includes/alert-use-https.md)]

    ```command&nbsp;line
    npm start
    ```

2. 「[テスト用に Outlook アドインをサイドロードする](../outlook/sideload-outlook-add-ins-for-testing.md)」の手順に従って Outlook アドインをサイドロードします。 [前のセクション](#configure-sso)の手順 3 で SSO を構成している間に Azure の接続に使用した Microsoft 365 管理者アカウントと同じ Microsoft 365 組織のメンバーであるユーザーで Outlook にサインインしている必要があります。 これにより、SSO を正常に実行するための適切な条件が確立されます。

3. Outlook で新しいメッセージを作成します。

4. メッセージ作成ウィンドウで、[ **タスクウィンドウの表示** ] ボタンを選択してアドイン作業ウィンドウを開きます。

    ![Outlook の [メッセージの作成] ウィンドウの [強調表示されたアドイン] リボン ボタンを示すスクリーン ショット。](../images/outlook-sso-ribbon-button.png)

5. 作業ウィンドウの下部にある [**マイ ユーザー プロファイルの情報を取得する**] ボタンを選択して、SSO プロセスを開始します。

6. アドインの代わりにアクセス許可を要求するダイアログ ウィンドウが表示される場合は、SSO はシナリオでサポートされず、代わりにアドインが別のユーザー認証方法に戻っていることを意味します。 これは、テナント管理者が Microsoft Graph にアクセスするためのアドインの同意を付与していない場合、またはユーザーが有効な Microsoft アカウントまたはMicrosoft 365 Educationまたは職場アカウントを使用して Office にサインインしていない場合に発生する可能性があります。 [ **承諾]** を選択して続行します。

    ![[承認] ボタンが強調表示された [アクセス許可] 要求ダイアログのスクリーンショット。](../images/sso-permissions-request.png)

    > [!NOTE]
    > ユーザーがこのアクセス許可の要求を受け入れると、今後再びプロンプトが表示されることはありません。

7. アドインは、サインインしたユーザーのプロファイル情報を取得し、メール メッセージの本文に書き込みます。

    ![Outlook の [メッセージの作成] ウィンドウのユーザー プロファイル情報を示すスクリーンショット。](../images/sso-user-profile-info-outlook.png)

## <a name="next-steps"></a>次の手順

おめでとうございます。可能な場合 SSO を使用し、SSO がサポートされていない場合は別のユーザー認証方法を使用する作業ウィンドウ アドインを正常に作成しました。 アドインをカスタマイズして異なる権限を必要とする新しい機能を追加する方法については、「[Node.js SSO が有効なアドインのカスタマイズする](sso-quickstart-customize.md)」をご覧ください。

## <a name="see-also"></a>関連項目

- [Office アドインのシングル サインオンを有効化する](../develop/sso-in-office-add-ins.md)
- [Node.js SSO が有効なアドインをカスタマイズする](sso-quickstart-customize.md)
- [シングル サインオンを使用する Node.js Office アドインを作成する](../develop/create-sso-office-add-ins-nodejs.md)
- [シングル サインオン (SSO) のエラー メッセージのトラブルシューティング](../develop/troubleshoot-sso-in-office-add-ins.md)
- [Visual Studio コードを使用して発行する](../publish/publish-add-in-vs-code.md#using-visual-studio-code-to-publish)