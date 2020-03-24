---
title: テスト用に Outlook アドインを展開してインストールする
description: マニフェスト ファイルを作成し、Web サーバーにアドイン UI ファイルを展開して、ユーザーのメールボックスにアドインをインストールします。その後、アドインをテストします。
ms.date: 03/18/2020
localization_priority: Priority
ms.openlocfilehash: 76688ad3e1eca2dda832a94c3a9ae815e37678bc
ms.sourcegitcommit: 6c381634c77d316f34747131860db0a0bced2529
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/21/2020
ms.locfileid: "42890978"
---
# <a name="deploy-and-install-outlook-add-ins-for-testing"></a><span data-ttu-id="e791b-103">テスト用に Outlook アドインを展開してインストールする</span><span class="sxs-lookup"><span data-stu-id="e791b-103">Deploy and install Outlook add-ins for testing</span></span>

<span data-ttu-id="e791b-104">Outlook アドインを開発するプロセスの一環として、テスト用にアドインの展開およびインストールを繰り返し行うことが多くあります。その場合は、以下の手順が必要です。</span><span class="sxs-lookup"><span data-stu-id="e791b-104">As part of the process of developing an Outlook add-in, you will probably find yourself iteratively deploying and installing the add-in for testing, which involves the following steps:</span></span>

1. <span data-ttu-id="e791b-105">アドインを記述したマニフェスト ファイルを作成します。</span><span class="sxs-lookup"><span data-stu-id="e791b-105">Creating a manifest file that describes the add-in.</span></span>
1. <span data-ttu-id="e791b-106">アドインの UI ファイルを Web サーバーに展開します。</span><span class="sxs-lookup"><span data-stu-id="e791b-106">Deploying the add-in UI file(s) to a web server.</span></span>
1. <span data-ttu-id="e791b-107">アドインをメールボックスにインストールします。</span><span class="sxs-lookup"><span data-stu-id="e791b-107">Installing the add-in in your mailbox.</span></span>
1. <span data-ttu-id="e791b-108">アドインをテストし、UI ファイルまたはマニフェスト ファイルを適切に変更します。さらに、手順 2 および 3 を繰り返して、変更箇所をテストします。</span><span class="sxs-lookup"><span data-stu-id="e791b-108">Testing the add-in, making appropriate changes to the UI or manifest files, and repeating steps 2 and 3 to test the changes.</span></span>

> [!NOTE]
> <span data-ttu-id="e791b-109">[カスタム ウィンドウは廃止された](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/)ため、[サポートされているアドイン拡張点](outlook-add-ins-overview.md#extension-points)を使用していることを確認してください。</span><span class="sxs-lookup"><span data-stu-id="e791b-109">[Custom panes have been deprecated](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/) so please ensure that you're using [a supported add-in extension point](outlook-add-ins-overview.md#extension-points).</span></span>

## <a name="create-a-manifest-file-for-the-add-in"></a><span data-ttu-id="e791b-110">アドイン用のマニフェスト ファイルを作成する</span><span class="sxs-lookup"><span data-stu-id="e791b-110">Create a manifest file for the add-in</span></span>

<span data-ttu-id="e791b-p101">各アドインは XML のマニフェストで記述されます。マニフェストは、アドインに関する情報をサーバーに提供し、ユーザーに向けたアドインについての説明的な情報を提供し、アドイン UI の HTML ファイルの場所を識別するドキュメントです。このマニフェストはローカル フォルダーにもサーバーにも保存できますが、その場所は、テストに使用するメールボックスの Exchange サーバーからアクセス可能な場所である必要があります。ここでの説明では、マニフェストがローカル フォルダーに保存されていることを想定しています。マニフェスト ファイルの作成方法については、「 [Outlook アドインのマニフェスト](manifests.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e791b-p101">Each add-in is described by an XML manifest, a document that gives the server information about the add-in, provides descriptive information about the add-in for the user, and identifies the location of the add-in UI HTML file. You can store the manifest in a local folder or server, as long as the location is accessible by the Exchange server of the mailbox that you are testing with. We'll assume that you store your manifest in a local folder. For information about how to create a manifest file, see [Outlook add-in manifests](manifests.md).</span></span>

## <a name="deploy-an-add-in-to-a-web-server"></a><span data-ttu-id="e791b-115">Web サーバーへのアドインを展開する</span><span class="sxs-lookup"><span data-stu-id="e791b-115">Deploy an add-in to a web server</span></span>

<span data-ttu-id="e791b-p102">HTML と JavaScript を使用してアドインを作成できます。作成されるソース ファイルは、アドインをホストする Exchange サーバーからアクセスできる Web サーバーに格納されます。アドインのソース ファイルを初期展開した後は、Web サーバー上に保存されている HTML ファイルまたは JavaScript ファイルを、新しいバージョンの HTML ファイルに置き換えることで、アドインの UI と動作を更新できます。</span><span class="sxs-lookup"><span data-stu-id="e791b-p102">You can use HTML and JavaScript to create the add-in. The resulting source files are stored on a web server that can be accessed by the Exchange server that hosts the add-in. After initially deploying the source files for the add-in, you can update the add-in UI and behavior by replacing the HTML files or JavaScript files stored on the web server with a new version of the HTML file.</span></span>

## <a name="install-the-add-in"></a><span data-ttu-id="e791b-119">アドインをインストールする</span><span class="sxs-lookup"><span data-stu-id="e791b-119">Install the add-in</span></span>

<span data-ttu-id="e791b-120">アドイン マニフェスト ファイルを準備して、アクセス可能な Web サーバーにアドイン UI を展開した後は、Outlook クライアントを使用するか、または Windows PowerShell コマンドレットをリモートで実行しアドインをインストールすることで、アドインを Exchange サーバーのメールボックスにサイドロードできます。</span><span class="sxs-lookup"><span data-stu-id="e791b-120">After preparing the add-in manifest file and deploying the add-in UI to a web server that can be accessed, you can sideload the add-in for a mailbox on an Exchange server by using an Outlook client, or install the add-in by running remote Windows PowerShell cmdlets.</span></span>

### <a name="sideload-the-add-in"></a><span data-ttu-id="e791b-121">アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="e791b-121">Sideload the add-in</span></span>

<span data-ttu-id="e791b-p103">メールボックスが Exchange Online、Exchange 2013 またはそれ以降のリリースのものである場合は、アドインをインストールできます。アドインをサイドロードするには、少なくとも Exchange Server の**自分のカスタム アプリ**の役割が必要です。アドイン マニフェストの URL またはファイル名を指定してアドインをテストしたり、一般的なアドインをインストールしたりする場合は、Exchange 管理者に連絡して、必要なアクセス許可を得る必要があります。</span><span class="sxs-lookup"><span data-stu-id="e791b-p103">You can install an add-in if your mailbox is on Exchange Online, Exchange 2013 or a later release. Sideloading add-ins requires at minimum the **My Custom Apps** role for your Exchange Server. In order to test your add-in, or install add-ins in general by specifying a URL or file name for the add-in manifest, you should request your Exchange administrator to provide the necessary permissions.</span></span>

<span data-ttu-id="e791b-p104">Exchange 管理者は、次のような PowerShell コマンドレットを実行して、必要なアクセス許可を単一ユーザーに割り当てることができます。この例では、`wendyri` は、ユーザーの電子メール エイリアスです。</span><span class="sxs-lookup"><span data-stu-id="e791b-p104">The Exchange administrator can run the following PowerShell cmdlet to assign a single user the necessary permissions. In this example, `wendyri` is the user's email alias.</span></span>

```powershell
New-ManagementRoleAssignment -Role "My Custom Apps" -User "wendyri"
```

<span data-ttu-id="e791b-127">必要な場合、管理者は次のようなコマンドレットを実行して、必要となる同様のアクセス許可を複数のユーザーに割り当てることができます。</span><span class="sxs-lookup"><span data-stu-id="e791b-127">If necessary, the administrator can run the following cmdlet to assign multiple users the similar necessary permissions:</span></span>

```powershell
$users = Get-Mailbox *$users | ForEach-Object { New-ManagementRoleAssignment -Role "My Custom Apps" -User $_.Alias}
```

<span data-ttu-id="e791b-128">自分のカスタム アドインの役割の詳細については、「["My Custom Apps/自分のカスタム アプリ" 役割](/exchange/my-custom-apps-role-exchange-2013-help)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="e791b-128">For more information about the My Custom Apps role, see [My Custom Apps role](/exchange/my-custom-apps-role-exchange-2013-help).</span></span>

<span data-ttu-id="e791b-129">Office 365 や Visual Studio を使用してアドインを開発すると、組織の管理者の役割が割り当てられ、EAC のファイルや URL を使用するか、Powershell コマンドレットを使用してアドインをインストールできるようになります。</span><span class="sxs-lookup"><span data-stu-id="e791b-129">Using Office 365 or Visual Studio to develop add-ins assigns you the organization administrator role which allows you to install add-ins by file or URL in the EAC, or by Powershell cmdlets.</span></span>

### <a name="install-an-add-in-by-using-remote-powershell"></a><span data-ttu-id="e791b-130">リモート PowerShell を使用してアドインをインストールする</span><span class="sxs-lookup"><span data-stu-id="e791b-130">Install an add-in by using remote PowerShell</span></span>

<span data-ttu-id="e791b-131">Exchange サーバー上に Windows PowerShell のリモート セッションを作成した後、次の PowerShell コマンドによって `New-App` コマンドレットを使用して Outlook アドインをインストールできます。</span><span class="sxs-lookup"><span data-stu-id="e791b-131">After you create a remote Windows PowerShell session on your Exchange server, you can install an Outlook add-in by using the `New-App` cmdlet with the following PowerShell command.</span></span>

```powershell
New-App -URL:"http://<fully-qualified URL">
```

<span data-ttu-id="e791b-132">完全修飾 URL は、アドイン用に準備したアドイン マニフェスト ファイルの場所です。</span><span class="sxs-lookup"><span data-stu-id="e791b-132">The fully qualified URL is the location of the add-in manifest file that you prepared for your add-in.</span></span>

<span data-ttu-id="e791b-133">さらに、次の PowerShell コマンドレットを使用すると、メールボックス用のアドインを管理できます。</span><span class="sxs-lookup"><span data-stu-id="e791b-133">You can use the following additional PowerShell cmdlets to manage the add-ins for a mailbox:</span></span>

-  <span data-ttu-id="e791b-134">`Get-App` - メールボックスに対して有効になっているアドインを一覧表示します。</span><span class="sxs-lookup"><span data-stu-id="e791b-134">`Get-App` - Lists the add-ins that are enabled for a mailbox.</span></span>
-  <span data-ttu-id="e791b-135">`Set-App` - メールボックスに対してアドインを有効または無効にします。</span><span class="sxs-lookup"><span data-stu-id="e791b-135">`Set-App` - Enables or disables a add-in on a mailbox.</span></span>
-  <span data-ttu-id="e791b-136">`Remove-App` - 現在インストールされているアドインを Exchange サーバーから削除します。</span><span class="sxs-lookup"><span data-stu-id="e791b-136">`Remove-App` - Removes a previously installed add-in from an Exchange server.</span></span>

## <a name="client-versions"></a><span data-ttu-id="e791b-137">クライアント バージョン</span><span class="sxs-lookup"><span data-stu-id="e791b-137">Client versions</span></span>

<span data-ttu-id="e791b-138">どのバージョンの Outlook クライアントをテストするかは、開発要件によって決まります。</span><span class="sxs-lookup"><span data-stu-id="e791b-138">Deciding what versions of the Outlook client to test depends on your development requirements.</span></span>

- <span data-ttu-id="e791b-p105">アドインを、個人用や組織のメンバー用に限って開発する場合は、自分の会社が使用している Outlook のバージョンをテストすることが重要です。一部のユーザーは Outlook on the web を使用する場合があるので、自分の会社で標準的に使用されているブラウザーのバージョンをテストすることも重要です。</span><span class="sxs-lookup"><span data-stu-id="e791b-p105">If you are developing an add-in for private use, or only for members of your organization, then it is important to test the versions of Outlook that your company uses. Keep in mind that some users may use Outlook on the web, so testing your company's standard browser versions is also important.</span></span>

- <span data-ttu-id="e791b-p106">[AppSource](https://appsource.microsoft.com) に一覧表示するアドインを開発する場合は、[Commercial marketplace の認定ポリシー 1120.3](/legal/marketplace/certification-policies#11203-functionality) で指定されている必要なバージョンをテストする必要があります。これには次が含まれます。</span><span class="sxs-lookup"><span data-stu-id="e791b-p106">If you are developing an add-in to list in [AppSource](https://appsource.microsoft.com), you must test the required versions as specified in the [Commercial marketplace certification policies 1120.3](/legal/marketplace/certification-policies#11203-functionality). This includes:</span></span>
    - <span data-ttu-id="e791b-143">Windows 用 Outlook の最新バージョンと最新の直前のバージョン。</span><span class="sxs-lookup"><span data-stu-id="e791b-143">The latest version of Outlook on Windows and the version prior to the latest.</span></span>
    - <span data-ttu-id="e791b-144">Mac 用 Outlook の最新バージョン。</span><span class="sxs-lookup"><span data-stu-id="e791b-144">The latest version of Outlook on Mac.</span></span>
    - <span data-ttu-id="e791b-145">iOS および Android 用の Outlook の最新バージョン (アドインが[モバイル フォーム ファクターをサポートしている](add-mobile-support.md)場合)。</span><span class="sxs-lookup"><span data-stu-id="e791b-145">The latest version of Outlook on iOS and Android (if your add-in [supports mobile form factor](add-mobile-support.md)).</span></span>
    - <span data-ttu-id="e791b-146">Commercial marketplace の検証ポリシー 1120.3 で指定されたブラウザーのバージョン。</span><span class="sxs-lookup"><span data-stu-id="e791b-146">The browser versions specified in the Commercial marketplace validation policy 1120.3.</span></span>

> [!NOTE]
> <span data-ttu-id="e791b-147">クライアントがサポートしていない [API 要件セットを要求しているために](apis.md)、アドインが上記のクライアントのいずれかをサポートしない場合は、そのクライアントが必要なクライアントのリストから削除されます。</span><span class="sxs-lookup"><span data-stu-id="e791b-147">If your add-in does not support one of the above clients due to [requesting an API requirement set](apis.md) that the client does not support, that client would be removed from the list of required clients.</span></span>

## <a name="see-also"></a><span data-ttu-id="e791b-148">関連項目</span><span class="sxs-lookup"><span data-stu-id="e791b-148">See also</span></span>

- [<span data-ttu-id="e791b-149">Office アドインでのユーザー エラーのトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="e791b-149">Troubleshoot user errors with Office Add-ins</span></span>](../testing/testing-and-troubleshooting.md)
