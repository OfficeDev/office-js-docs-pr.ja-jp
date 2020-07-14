---
title: テスト用に Outlook アドインを展開してインストールする
description: マニフェスト ファイルを作成し、Web サーバーにアドイン UI ファイルを展開して、ユーザーのメールボックスにアドインをインストールします。その後、アドインをテストします。
ms.date: 05/20/2020
localization_priority: Priority
ms.openlocfilehash: 97841f7c8112b42cee2927f238b31fe985b2e101
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093862"
---
# <a name="deploy-and-install-outlook-add-ins-for-testing"></a><span data-ttu-id="cc7b0-103">テスト用に Outlook アドインを展開してインストールする</span><span class="sxs-lookup"><span data-stu-id="cc7b0-103">Deploy and install Outlook add-ins for testing</span></span>

<span data-ttu-id="cc7b0-104">Outlook アドインを開発するプロセスの一環として、テスト用にアドインの展開およびインストールを繰り返し行うことが多くあります。その場合は、以下の手順が必要です。</span><span class="sxs-lookup"><span data-stu-id="cc7b0-104">As part of the process of developing an Outlook add-in, you will probably find yourself iteratively deploying and installing the add-in for testing, which involves the following steps:</span></span>

1. <span data-ttu-id="cc7b0-105">アドインを記述したマニフェスト ファイルを作成します。</span><span class="sxs-lookup"><span data-stu-id="cc7b0-105">Creating a manifest file that describes the add-in.</span></span>
1. <span data-ttu-id="cc7b0-106">アドインの UI ファイルを Web サーバーに展開します。</span><span class="sxs-lookup"><span data-stu-id="cc7b0-106">Deploying the add-in UI file(s) to a web server.</span></span>
1. <span data-ttu-id="cc7b0-107">アドインをメールボックスにインストールします。</span><span class="sxs-lookup"><span data-stu-id="cc7b0-107">Installing the add-in in your mailbox.</span></span>
1. <span data-ttu-id="cc7b0-108">アドインをテストし、UI ファイルまたはマニフェスト ファイルを適切に変更します。さらに、手順 2 および 3 を繰り返して、変更箇所をテストします。</span><span class="sxs-lookup"><span data-stu-id="cc7b0-108">Testing the add-in, making appropriate changes to the UI or manifest files, and repeating steps 2 and 3 to test the changes.</span></span>

> [!NOTE]
> <span data-ttu-id="cc7b0-109">[カスタム ウィンドウは廃止された](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/)ため、[サポートされているアドイン拡張点](outlook-add-ins-overview.md#extension-points)を使用していることを確認してください。</span><span class="sxs-lookup"><span data-stu-id="cc7b0-109">[Custom panes have been deprecated](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/) so please ensure that you're using [a supported add-in extension point](outlook-add-ins-overview.md#extension-points).</span></span>

## <a name="create-a-manifest-file-for-the-add-in"></a><span data-ttu-id="cc7b0-110">アドイン用のマニフェスト ファイルを作成する</span><span class="sxs-lookup"><span data-stu-id="cc7b0-110">Create a manifest file for the add-in</span></span>

<span data-ttu-id="cc7b0-111">Each add-in is described by an XML manifest, a document that gives the server information about the add-in, provides descriptive information about the add-in for the user, and identifies the location of the add-in UI HTML file.</span><span class="sxs-lookup"><span data-stu-id="cc7b0-111">Each add-in is described by an XML manifest, a document that gives the server information about the add-in, provides descriptive information about the add-in for the user, and identifies the location of the add-in UI HTML file.</span></span> <span data-ttu-id="cc7b0-112">You can store the manifest in a local folder or server, as long as the location is accessible by the Exchange server of the mailbox that you are testing with.</span><span class="sxs-lookup"><span data-stu-id="cc7b0-112">You can store the manifest in a local folder or server, as long as the location is accessible by the Exchange server of the mailbox that you are testing with.</span></span> <span data-ttu-id="cc7b0-113">We'll assume that you store your manifest in a local folder.</span><span class="sxs-lookup"><span data-stu-id="cc7b0-113">We'll assume that you store your manifest in a local folder.</span></span> <span data-ttu-id="cc7b0-114">For information about how to create a manifest file, see [Outlook add-in manifests](manifests.md).</span><span class="sxs-lookup"><span data-stu-id="cc7b0-114">For information about how to create a manifest file, see [Outlook add-in manifests](manifests.md).</span></span>

## <a name="deploy-an-add-in-to-a-web-server"></a><span data-ttu-id="cc7b0-115">Web サーバーへのアドインを展開する</span><span class="sxs-lookup"><span data-stu-id="cc7b0-115">Deploy an add-in to a web server</span></span>

<span data-ttu-id="cc7b0-116">You can use HTML and JavaScript to create the add-in.</span><span class="sxs-lookup"><span data-stu-id="cc7b0-116">You can use HTML and JavaScript to create the add-in.</span></span> <span data-ttu-id="cc7b0-117">The resulting source files are stored on a web server that can be accessed by the Exchange server that hosts the add-in.</span><span class="sxs-lookup"><span data-stu-id="cc7b0-117">The resulting source files are stored on a web server that can be accessed by the Exchange server that hosts the add-in.</span></span> <span data-ttu-id="cc7b0-118">After initially deploying the source files for the add-in, you can update the add-in UI and behavior by replacing the HTML files or JavaScript files stored on the web server with a new version of the HTML file.</span><span class="sxs-lookup"><span data-stu-id="cc7b0-118">After initially deploying the source files for the add-in, you can update the add-in UI and behavior by replacing the HTML files or JavaScript files stored on the web server with a new version of the HTML file.</span></span>

## <a name="install-the-add-in"></a><span data-ttu-id="cc7b0-119">アドインをインストールする</span><span class="sxs-lookup"><span data-stu-id="cc7b0-119">Install the add-in</span></span>

<span data-ttu-id="cc7b0-120">アドイン マニフェスト ファイルを準備して、アクセス可能な Web サーバーにアドイン UI を展開した後は、Outlook クライアントを使用するか、または Windows PowerShell コマンドレットをリモートで実行しアドインをインストールすることで、アドインを Exchange サーバーのメールボックスにサイドロードできます。</span><span class="sxs-lookup"><span data-stu-id="cc7b0-120">After preparing the add-in manifest file and deploying the add-in UI to a web server that can be accessed, you can sideload the add-in for a mailbox on an Exchange server by using an Outlook client, or install the add-in by running remote Windows PowerShell cmdlets.</span></span>

### <a name="sideload-the-add-in"></a><span data-ttu-id="cc7b0-121">アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="cc7b0-121">Sideload the add-in</span></span>

<span data-ttu-id="cc7b0-122">You can install an add-in if your mailbox is on Exchange Online, Exchange 2013 or a later release.</span><span class="sxs-lookup"><span data-stu-id="cc7b0-122">You can install an add-in if your mailbox is on Exchange Online, Exchange 2013 or a later release.</span></span> <span data-ttu-id="cc7b0-123">Sideloading add-ins requires at minimum the **My Custom Apps** role for your Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="cc7b0-123">Sideloading add-ins requires at minimum the **My Custom Apps** role for your Exchange Server.</span></span> <span data-ttu-id="cc7b0-124">In order to test your add-in, or install add-ins in general by specifying a URL or file name for the add-in manifest, you should request your Exchange administrator to provide the necessary permissions.</span><span class="sxs-lookup"><span data-stu-id="cc7b0-124">In order to test your add-in, or install add-ins in general by specifying a URL or file name for the add-in manifest, you should request your Exchange administrator to provide the necessary permissions.</span></span>

<span data-ttu-id="cc7b0-125">The Exchange administrator can run the following PowerShell cmdlet to assign a single user the necessary permissions.</span><span class="sxs-lookup"><span data-stu-id="cc7b0-125">The Exchange administrator can run the following PowerShell cmdlet to assign a single user the necessary permissions.</span></span> <span data-ttu-id="cc7b0-126">In this example, `wendyri` is the user's email alias.</span><span class="sxs-lookup"><span data-stu-id="cc7b0-126">In this example, `wendyri` is the user's email alias.</span></span>

```powershell
New-ManagementRoleAssignment -Role "My Custom Apps" -User "wendyri"
```

<span data-ttu-id="cc7b0-127">必要な場合、管理者は次のようなコマンドレットを実行して、必要となる同様のアクセス許可を複数のユーザーに割り当てることができます。</span><span class="sxs-lookup"><span data-stu-id="cc7b0-127">If necessary, the administrator can run the following cmdlet to assign multiple users the similar necessary permissions:</span></span>

```powershell
$users = Get-Mailbox *$users | ForEach-Object { New-ManagementRoleAssignment -Role "My Custom Apps" -User $_.Alias}
```

<span data-ttu-id="cc7b0-128">自分のカスタム アドインの役割の詳細については、「["My Custom Apps/自分のカスタム アプリ" 役割](/exchange/my-custom-apps-role-exchange-2013-help)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="cc7b0-128">For more information about the My Custom Apps role, see [My Custom Apps role](/exchange/my-custom-apps-role-exchange-2013-help).</span></span>

<span data-ttu-id="cc7b0-129">Microsoft 365 や Visual Studio を使用してアドインを開発すると、組織の管理者の役割が割り当てられ、EAC のファイルや URL を使用するか、Powershell コマンドレットを使用してアドインをインストールできるようになります。</span><span class="sxs-lookup"><span data-stu-id="cc7b0-129">Using Microsoft 365 or Visual Studio to develop add-ins assigns you the organization administrator role which allows you to install add-ins by file or URL in the EAC, or by Powershell cmdlets.</span></span>

### <a name="install-an-add-in-by-using-remote-powershell"></a><span data-ttu-id="cc7b0-130">リモート PowerShell を使用してアドインをインストールする</span><span class="sxs-lookup"><span data-stu-id="cc7b0-130">Install an add-in by using remote PowerShell</span></span>

<span data-ttu-id="cc7b0-131">Exchange サーバー上に Windows PowerShell のリモート セッションを作成した後、次の PowerShell コマンドによって `New-App` コマンドレットを使用して Outlook アドインをインストールできます。</span><span class="sxs-lookup"><span data-stu-id="cc7b0-131">After you create a remote Windows PowerShell session on your Exchange server, you can install an Outlook add-in by using the `New-App` cmdlet with the following PowerShell command.</span></span>

```powershell
New-App -URL:"http://<fully-qualified URL">
```

<span data-ttu-id="cc7b0-132">完全修飾 URL は、アドイン用に準備したアドイン マニフェスト ファイルの場所です。</span><span class="sxs-lookup"><span data-stu-id="cc7b0-132">The fully qualified URL is the location of the add-in manifest file that you prepared for your add-in.</span></span>

<span data-ttu-id="cc7b0-133">さらに、次の PowerShell コマンドレットを使用すると、メールボックス用のアドインを管理できます。</span><span class="sxs-lookup"><span data-stu-id="cc7b0-133">You can use the following additional PowerShell cmdlets to manage the add-ins for a mailbox:</span></span>

- <span data-ttu-id="cc7b0-134">`Get-App` - メールボックスに対して有効になっているアドインを一覧表示します。</span><span class="sxs-lookup"><span data-stu-id="cc7b0-134">`Get-App` - Lists the add-ins that are enabled for a mailbox.</span></span>
- <span data-ttu-id="cc7b0-135">`Set-App` - メールボックスに対してアドインを有効または無効にします。</span><span class="sxs-lookup"><span data-stu-id="cc7b0-135">`Set-App` - Enables or disables a add-in on a mailbox.</span></span>
- <span data-ttu-id="cc7b0-136">`Remove-App` - 現在インストールされているアドインを Exchange サーバーから削除します。</span><span class="sxs-lookup"><span data-stu-id="cc7b0-136">`Remove-App` - Removes a previously installed add-in from an Exchange server.</span></span>

## <a name="client-versions"></a><span data-ttu-id="cc7b0-137">クライアント バージョン</span><span class="sxs-lookup"><span data-stu-id="cc7b0-137">Client versions</span></span>

<span data-ttu-id="cc7b0-138">どのバージョンの Outlook クライアントをテストするかは、開発要件によって決まります。</span><span class="sxs-lookup"><span data-stu-id="cc7b0-138">Deciding what versions of the Outlook client to test depends on your development requirements.</span></span>

- <span data-ttu-id="cc7b0-139">If you are developing an add-in for private use, or only for members of your organization, then it is important to test the versions of Outlook that your company uses.</span><span class="sxs-lookup"><span data-stu-id="cc7b0-139">If you are developing an add-in for private use, or only for members of your organization, then it is important to test the versions of Outlook that your company uses.</span></span> <span data-ttu-id="cc7b0-140">Keep in mind that some users may use Outlook on the web, so testing your company's standard browser versions is also important.</span><span class="sxs-lookup"><span data-stu-id="cc7b0-140">Keep in mind that some users may use Outlook on the web, so testing your company's standard browser versions is also important.</span></span>

- <span data-ttu-id="cc7b0-141">If you are developing an add-in to list in [AppSource](https://appsource.microsoft.com), you must test the required versions as specified in the [Commercial marketplace certification policies 1120.3](/legal/marketplace/certification-policies#11203-functionality).</span><span class="sxs-lookup"><span data-stu-id="cc7b0-141">If you are developing an add-in to list in [AppSource](https://appsource.microsoft.com), you must test the required versions as specified in the [Commercial marketplace certification policies 1120.3](/legal/marketplace/certification-policies#11203-functionality).</span></span> <span data-ttu-id="cc7b0-142">This includes:</span><span class="sxs-lookup"><span data-stu-id="cc7b0-142">This includes:</span></span>
  - <span data-ttu-id="cc7b0-143">Windows 用 Outlook の最新バージョンと最新の直前のバージョン。</span><span class="sxs-lookup"><span data-stu-id="cc7b0-143">The latest version of Outlook on Windows and the version prior to the latest.</span></span>
  - <span data-ttu-id="cc7b0-144">Mac 用 Outlook の最新バージョン。</span><span class="sxs-lookup"><span data-stu-id="cc7b0-144">The latest version of Outlook on Mac.</span></span>
  - <span data-ttu-id="cc7b0-145">iOS および Android 用の Outlook の最新バージョン (アドインが[モバイル フォーム ファクターをサポートしている](add-mobile-support.md)場合)。</span><span class="sxs-lookup"><span data-stu-id="cc7b0-145">The latest version of Outlook on iOS and Android (if your add-in [supports mobile form factor](add-mobile-support.md)).</span></span>
  - <span data-ttu-id="cc7b0-146">Commercial marketplace の検証ポリシー 1120.3 で指定されたブラウザーのバージョン。</span><span class="sxs-lookup"><span data-stu-id="cc7b0-146">The browser versions specified in the Commercial marketplace validation policy 1120.3.</span></span>

> [!NOTE]
> <span data-ttu-id="cc7b0-147">クライアントがサポートしていない [API 要件セットを要求しているために](apis.md)、アドインが上記のクライアントのいずれかをサポートしない場合は、そのクライアントが必要なクライアントのリストから削除されます。</span><span class="sxs-lookup"><span data-stu-id="cc7b0-147">If your add-in does not support one of the above clients due to [requesting an API requirement set](apis.md) that the client does not support, that client would be removed from the list of required clients.</span></span>

## <a name="outlook-on-the-web-and-exchange-server-versions"></a><span data-ttu-id="cc7b0-148">Outlook on the web および Exchange サーバーのバージョン</span><span class="sxs-lookup"><span data-stu-id="cc7b0-148">Outlook on the web and Exchange server versions</span></span>

<span data-ttu-id="cc7b0-149">顧客および Microsoft 365 アカウントのユーザーは、Outlook on the web にアクセスすると最新の UI バージョンを表示し、廃止されたクラシック バージョンを表示しなくなります。</span><span class="sxs-lookup"><span data-stu-id="cc7b0-149">Consumer and Microsoft 365 account users see the modern UI version when they access Outlook on the web and no longer see the classic version which has been deprecated.</span></span> <span data-ttu-id="cc7b0-150">ただし、オンプレミスの Exchange サーバーは、従来の Outlook on the web を引き続きサポートします。</span><span class="sxs-lookup"><span data-stu-id="cc7b0-150">However, on-premises Exchange servers continue to support classic Outlook on the web.</span></span> <span data-ttu-id="cc7b0-151">したがって、検証プロセス中に、提出物はアドインが従来の Outlook on the web と互換性がないという警告を受け取る場合があります。</span><span class="sxs-lookup"><span data-stu-id="cc7b0-151">Therefore, during the validation process, your submission may receive a warning that the add-in is not compatible with classic Outlook on the web.</span></span> <span data-ttu-id="cc7b0-152">その場合は、オンプレミスの Exchange 環境でアドインをテストすることを検討する必要があります。</span><span class="sxs-lookup"><span data-stu-id="cc7b0-152">In that case, you should consider testing your add-in in an on-premises Exchange environment.</span></span> <span data-ttu-id="cc7b0-153">この警告によって AppSource への送信がブロックされることはありませんが、顧客がオンプレミスの Exchange 環境で Outlook on the web を使用すると、次善のエクスペリエンスが発生する可能性があります。</span><span class="sxs-lookup"><span data-stu-id="cc7b0-153">This warning won't block your submission to AppSource but your customers may experience a sub-optimal experience if they use Outlook on the web in an on-premises Exchange environment.</span></span>

<span data-ttu-id="cc7b0-154">これを軽減するために、独自のプライベート オンプレミス Exchange 環境に接続された Outlook on the web でアドインをテストすることをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="cc7b0-154">To mitigate this, we recommend you test your add-in in Outlook on the web connected to your own private on-premises Exchange environment.</span></span> <span data-ttu-id="cc7b0-155">詳細については、[Exchange 2016 または Exchange 2019 テスト環境を確立する](/Exchange/plan-and-deploy/plan-and-deploy?view=exchserver-2019#establish-an-exchange-2016-or-exchange-2019-test-environment)方法と、[Exchange Server で Outlook on the web](/exchange/clients/outlook-on-the-web/outlook-on-the-web?view=exchserver-2019) を管理する方法に関するガイダンスを参照してください。</span><span class="sxs-lookup"><span data-stu-id="cc7b0-155">For more information, see guidance on how to [Establish an Exchange 2016 or Exchange 2019 test environment](/Exchange/plan-and-deploy/plan-and-deploy?view=exchserver-2019#establish-an-exchange-2016-or-exchange-2019-test-environment) and how to manage [Outlook on the web in Exchange Server](/exchange/clients/outlook-on-the-web/outlook-on-the-web?view=exchserver-2019).</span></span>

<span data-ttu-id="cc7b0-156">または、オンプレミスの Exchange サーバーをホストおよび管理するサービスの料金を支払い、使用することもできます。</span><span class="sxs-lookup"><span data-stu-id="cc7b0-156">Alternatively, you can opt to pay for and use a service that hosts and manages on-premises Exchange servers.</span></span> <span data-ttu-id="cc7b0-157">いくつかのオプションがあります:</span><span class="sxs-lookup"><span data-stu-id="cc7b0-157">A couple of options are:</span></span>

- [<span data-ttu-id="cc7b0-158">Rackspace</span><span class="sxs-lookup"><span data-stu-id="cc7b0-158">Rackspace</span></span>](https://www.rackspace.com/email-hosting/exchange-server)
- [<span data-ttu-id="cc7b0-159">Hostway</span><span class="sxs-lookup"><span data-stu-id="cc7b0-159">Hostway</span></span>](https://hostway.com/products-services-2/hosted-microsoft-exchange/)

<span data-ttu-id="cc7b0-160">さらに、オンプレミスの Exchange に接続しているユーザーがアドインを使用できないようにする場合は、アドイン マニフェストの[要件セット](../reference/requirement-sets/outlook-api-requirement-sets.md#exchange-server-support)を 1.6 以上に設定できます。</span><span class="sxs-lookup"><span data-stu-id="cc7b0-160">Furthermore, if you don't want your add-ins to be available for users who are connected to on-premises Exchange, you can set the [requirement set](../reference/requirement-sets/outlook-api-requirement-sets.md#exchange-server-support) in the add-in manifest to be 1.6 or higher.</span></span> <span data-ttu-id="cc7b0-161">このようなアドインは、従来の Outlook on the Web UI ではテストまたは検証されません。</span><span class="sxs-lookup"><span data-stu-id="cc7b0-161">Such add-ins will not be tested or validated on the classic Outlook on the web UI.</span></span>

## <a name="see-also"></a><span data-ttu-id="cc7b0-162">関連項目</span><span class="sxs-lookup"><span data-stu-id="cc7b0-162">See also</span></span>

- [<span data-ttu-id="cc7b0-163">Office アドインでのユーザー エラーのトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="cc7b0-163">Troubleshoot user errors with Office Add-ins</span></span>](../testing/testing-and-troubleshooting.md)
