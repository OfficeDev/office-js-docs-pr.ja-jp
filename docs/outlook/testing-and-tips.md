---
title: テスト用に Outlook アドインを展開してインストールする
description: マニフェスト ファイルを作成し、Web サーバーにアドイン UI ファイルを展開して、ユーザーのメールボックスにアドインをインストールします。その後、アドインをテストします。
ms.date: 10/18/2022
ms.localizationpriority: high
ms.openlocfilehash: 1b6d29fa85b855adbf75a33345850582d2eecc02
ms.sourcegitcommit: eca6c16d0bb74bed2d35a21723dd98c6b41ef507
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/18/2022
ms.locfileid: "68607521"
---
# <a name="deploy-and-install-outlook-add-ins-for-testing"></a>テスト用に Outlook アドインを展開してインストールする

Outlook アドインを開発するプロセスの一環として、テスト用にアドインの展開およびインストールを繰り返し行うことが多くあります。その場合は、以下の手順が必要です。

1. アドインを記述したマニフェスト ファイルを作成します。
1. アドインの UI ファイルを Web サーバーに展開します。
1. アドインをメールボックスにインストールします。
1. アドインをテストし、UI ファイルまたはマニフェスト ファイルを適切に変更します。さらに、手順 2 および 3 を繰り返して、変更箇所をテストします。

> [!NOTE]
> [カスタム ウィンドウは廃止された](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/)ため、[サポートされているアドイン拡張点](outlook-add-ins-overview.md#extension-points)を使用していることを確認してください。

## <a name="create-a-manifest-file-for-the-add-in"></a>アドイン用のマニフェスト ファイルを作成する

各アドインはマニフェスト (アドインに関するサーバー情報を提供するドキュメント) によって記述され、ユーザーのアドインに関するわかりやすい情報を提供し、アドイン UI HTML ファイルの場所を識別します。 テストで使用するメールボックスの Exchange サーバーがアクセスできるローカル フォルダーまたはサーバーである限り、マニフェストはどの場所にでも格納できます。 ここでは、マニフェストをローカル フォルダーに格納することを前提とします。 マニフェスト ファイルを作成する方法については、「[Outlook アドインのマニフェスト](manifests.md)」をご覧ください。

## <a name="deploy-an-add-in-to-a-web-server"></a>Web サーバーへのアドインを展開する

You can use HTML and JavaScript to create the add-in. The resulting source files are stored on a web server that can be accessed by the Exchange server that hosts the add-in. After initially deploying the source files for the add-in, you can update the add-in UI and behavior by replacing the HTML files or JavaScript files stored on the web server with a new version of the HTML file.

## <a name="install-the-add-in"></a>アドインをインストールする

アドイン マニフェスト ファイルを準備して、アクセス可能な Web サーバーにアドイン UI を展開した後は、Outlook クライアントを使用するか、または Windows PowerShell コマンドレットをリモートで実行しアドインをインストールすることで、アドインを Exchange サーバーのメールボックスにサイドロードできます。

### <a name="sideload-the-add-in"></a>アドインをサイドロードする

You can install an add-in if your mailbox is on Exchange Online, Exchange 2013 or a later release. Sideloading add-ins requires at minimum the **My Custom Apps** role for your Exchange Server. In order to test your add-in, or install add-ins in general by specifying a URL or file name for the add-in manifest, you should request your Exchange administrator to provide the necessary permissions.

The Exchange administrator can run the following PowerShell cmdlet to assign a single user the necessary permissions. In this example, `wendyri` is the user's email alias.

```powershell
New-ManagementRoleAssignment -Role "My Custom Apps" -User "wendyri"
```

必要な場合、管理者は次のようなコマンドレットを実行して、必要となる同様のアクセス許可を複数のユーザーに割り当てることができます。

```powershell
$users = Get-Mailbox *$users | ForEach-Object { New-ManagementRoleAssignment -Role "My Custom Apps" -User $_.Alias}
```

自分のカスタム アドインの役割の詳細については、「["My Custom Apps/自分のカスタム アプリ" 役割](/exchange/my-custom-apps-role-exchange-2013-help)」をご覧ください。

Microsoft 365 や Visual Studio を使用してアドインを開発すると、組織の管理者の役割が割り当てられ、EAC のファイルや URL を使用するか、Powershell コマンドレットを使用してアドインをインストールできるようになります。

### <a name="install-an-add-in-by-using-remote-powershell"></a>リモート PowerShell を使用してアドインをインストールする

Exchange サーバー上に Windows PowerShell のリモート セッションを作成した後、次の PowerShell コマンドによって `New-App` コマンドレットを使用して Outlook アドインをインストールできます。

```powershell
New-App -URL:"http://<fully-qualified URL">
```

完全修飾 URL は、アドイン用に準備したアドイン マニフェスト ファイルの場所です。

さらに、次の PowerShell コマンドレットを使用して、メールボックス用のアドインを管理します。

- `Get-App` - メールボックスに対して有効になっているアドインを一覧表示します。
- `Set-App` - メールボックスに対してアドインを有効または無効にします。
- `Remove-App` - 現在インストールされているアドインを Exchange サーバーから削除します。

## <a name="client-versions"></a>クライアント バージョン

どのバージョンの Outlook クライアントをテストするかは、開発要件によって決まります。

- If you're developing an add-in for private use, or only for members of your organization, then it is important to test the versions of Outlook that your company uses. Keep in mind that some users may use Outlook on the web, so testing your company's standard browser versions is also important.

- If you're developing an add-in to list in [AppSource](https://appsource.microsoft.com), you must test the required versions as specified in the [Commercial marketplace certification policies 1120.3](/legal/marketplace/certification-policies#11203-functionality). This includes:
  - Windows 用 Outlook の最新バージョンと最新の直前のバージョン。
  - Mac 用 Outlook の最新バージョン。
  - iOS および Android 用の Outlook の最新バージョン (アドインが[モバイル フォーム ファクターをサポートしている](add-mobile-support.md)場合)。
  - Commercial marketplace の検証ポリシー 1120.3 で指定されたブラウザーのバージョン。

> [!NOTE]
> クライアントがサポートしていない [API 要件セットを要求しているために](apis.md)、アドインが上記のクライアントのいずれかをサポートしない場合は、そのクライアントが必要なクライアントのリストから削除されます。

## <a name="outlook-on-the-web-and-exchange-server-versions"></a>Outlook on the web および Exchange サーバーのバージョン

顧客および Microsoft 365 アカウントのユーザーは、Outlook on the web にアクセスすると最新の UI バージョンを表示し、廃止されたクラシック バージョンを表示しなくなります。 ただし、オンプレミスの Exchange サーバーは、従来の Outlook on the web を引き続きサポートします。 したがって、検証プロセス中に、提出物はアドインが従来の Outlook on the web と互換性がないという警告を受け取る場合があります。 その場合は、オンプレミスの Exchange 環境でアドインをテストすることを検討する必要があります。 この警告によって AppSource への送信がブロックされることはありませんが、顧客がオンプレミスの Exchange 環境で Outlook on the web を使用すると、次善のエクスペリエンスが発生する可能性があります。

これを軽減するために、独自のプライベート オンプレミス Exchange 環境に接続された Outlook on the web でアドインをテストすることをお勧めします。 詳細については、[Exchange 2016 または Exchange 2019 テスト環境を確立する](/Exchange/plan-and-deploy/plan-and-deploy?view=exchserver-2019&preserve-view=true#establish-an-exchange-2016-or-exchange-2019-test-environment)方法と、[Exchange Server で Outlook on the web](/exchange/clients/outlook-on-the-web/outlook-on-the-web?view=exchserver-2019&preserve-view=true) を管理する方法に関するガイダンスを参照してください。

または、オンプレミスの Exchange サーバーをホストおよび管理するサービスの料金を支払い、使用することもできます。 いくつかのオプションがあります:

- [Rackspace](https://www.rackspace.com/email-hosting/exchange-server)
- [Hostway](https://hostway.com/microsoft-exchange/)

さらに、オンプレミスの Exchange に接続しているユーザーがアドインを使用できないようにする場合は、アドイン マニフェストの[要件セット](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#exchange-server-support)を 1.6 以上に設定できます。 このようなアドインは、従来の Outlook on the Web UI ではテストまたは検証されません。

## <a name="see-also"></a>関連項目

- [Office アドインでのユーザー エラーのトラブルシューティング](../testing/testing-and-troubleshooting.md)
