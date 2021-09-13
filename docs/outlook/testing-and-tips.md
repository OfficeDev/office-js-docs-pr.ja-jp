---
title: テスト用に Outlook アドインを展開してインストールする
description: マニフェスト ファイルを作成し、Web サーバーにアドイン UI ファイルを展開して、ユーザーのメールボックスにアドインをインストールします。その後、アドインをテストします。
ms.date: 07/08/2021
ms.localizationpriority: high
ms.openlocfilehash: f9c9fa23dc59a36bee608fe8115b3ad77cfab623
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2021
ms.locfileid: "59151381"
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

各アドインは、XML マニフェストによって記述されます。XML マニフェストは、サーバーに対してアドインに関する情報を提供し、ユーザーに対してアドインについて説明するドキュメントです。また、アドイン UI の HTML ファイルの場所も XML マニフェストで指定されます。 テストで使用するメールボックスの Exchange サーバーがアクセスできるローカル フォルダーまたはサーバーである限り、マニフェストはどの場所にでも格納できます。 ここでは、マニフェストをローカル フォルダーに格納することを前提とします。 マニフェスト ファイルを作成する方法については、「[Outlook アドインのマニフェスト](manifests.md)」をご覧ください。

## <a name="deploy-an-add-in-to-a-web-server"></a>Web サーバーへのアドインを展開する

HTML と JavaScript を使用してアドインを作成できます。作成されるソース ファイルは、アドインをホストする Exchange サーバーからアクセスできる Web サーバーに格納されます。アドインのソース ファイルを初期展開した後は、Web サーバー上に保存されている HTML ファイルまたは JavaScript ファイルを、新しいバージョンの HTML ファイルに置き換えることで、アドインの UI と動作を更新できます。

## <a name="install-the-add-in"></a>アドインをインストールする

アドイン マニフェスト ファイルを準備して、アクセス可能な Web サーバーにアドイン UI を展開した後は、Outlook クライアントを使用するか、または Windows PowerShell コマンドレットをリモートで実行しアドインをインストールすることで、アドインを Exchange サーバーのメールボックスにサイドロードできます。

### <a name="sideload-the-add-in"></a>アドインをサイドロードする

メールボックスが Exchange Online、Exchange 2013 またはそれ以降のリリースのものである場合は、アドインをインストールできます。アドインをサイドロードするには、少なくとも Exchange Server の **自分のカスタム アプリ** の役割が必要です。アドイン マニフェストの URL またはファイル名を指定してアドインをテストしたり、一般的なアドインをインストールしたりする場合は、Exchange 管理者に連絡して、必要なアクセス許可を得る必要があります。

Exchange 管理者は、次のような PowerShell コマンドレットを実行して、必要なアクセス許可を単一ユーザーに割り当てることができます。この例では、`wendyri` は、ユーザーの電子メール エイリアスです。

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

- アドインを、個人用や組織のメンバー用に限って開発する場合は、自分の会社が使用している Outlook のバージョンをテストすることが重要です。一部のユーザーは Outlook on the web を使用する場合があるので、自分の会社で標準的に使用されているブラウザーのバージョンをテストすることも重要です。

- [AppSource](https://appsource.microsoft.com) に一覧表示するアドインを開発する場合は、[Commercial marketplace の認定ポリシー 1120.3](/legal/marketplace/certification-policies#11203-functionality) で指定されている必要なバージョンをテストする必要があります。 これには、次のものが含まれます。
  - Windows 用 Outlook の最新バージョンと最新の直前のバージョン。
  - Mac 用 Outlook の最新バージョン。
  - iOS および Android 用の Outlook の最新バージョン (アドインが[モバイル フォーム ファクターをサポートしている](add-mobile-support.md)場合)。
  - Commercial marketplace の検証ポリシー 1120.3 で指定されたブラウザーのバージョン。

> [!NOTE]
> クライアントがサポートしていない [API 要件セットを要求しているために](apis.md)、アドインが上記のクライアントのいずれかをサポートしない場合は、そのクライアントが必要なクライアントのリストから削除されます。

## <a name="outlook-on-the-web-and-exchange-server-versions"></a>Outlook on the web および Exchange サーバーのバージョン

顧客および Microsoft 365 アカウントのユーザーは、Outlook on the web にアクセスすると最新の UI バージョンを表示し、廃止されたクラシック バージョンを表示しなくなります。 ただし、オンプレミスの Exchange サーバーは、従来の Outlook on the web を引き続きサポートします。 したがって、検証プロセス中に、提出物はアドインが従来の Outlook on the web と互換性がないという警告を受け取る場合があります。 その場合は、オンプレミスの Exchange 環境でアドインをテストすることを検討する必要があります。 この警告によって AppSource への送信がブロックされることはありませんが、顧客がオンプレミスの Exchange 環境で Outlook on the web を使用すると、次善のエクスペリエンスが発生する可能性があります。

これを軽減するために、独自のプライベート オンプレミス Exchange 環境に接続された Outlook on the web でアドインをテストすることをお勧めします。 詳細については、[Exchange 2016 または Exchange 2019 テスト環境を確立する](/Exchange/plan-and-deploy/plan-and-deploy?view=exchserver-2019&preserve-view=true#establish-an-exchange-2016-or-exchange-2019-test-environment)方法と、[Exchange Server で Outlook on the web](/exchange/clients/outlook-on-the-web/outlook-on-the-web?view=exchserver-2019&preserve-view=true) を管理する方法に関するガイダンスを参照してください。

または、オンプレミスの Exchange サーバーをホストおよび管理するサービスの料金を支払い、使用することもできます。いくつかのオプションは次のとおりです。

- [Rackspace](https://www.rackspace.com/email-hosting/exchange-server)
- [Hostway](https://hostway.com/microsoft-exchange/)

さらに、オンプレミスの Exchange に接続しているユーザーがアドインを使用できないようにする場合は、アドイン マニフェストの[要件セット](../reference/requirement-sets/outlook-api-requirement-sets.md#exchange-server-support)を 1.6 以上に設定できます。 このようなアドインは、従来の Outlook on the Web UI ではテストまたは検証されません。

## <a name="see-also"></a>関連項目

- [Office アドインでのユーザー エラーのトラブルシューティング](../testing/testing-and-troubleshooting.md)
