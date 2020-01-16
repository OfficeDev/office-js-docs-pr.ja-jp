---
title: Yeoman ジェネレーターを使用して、SSO を使用する Office アドインを作成する (プレビュー)
description: Yeoman ジェネレーターを使用して、シングル サインオンを使用する Node.js Office アドインを作成する (プレビュー)
ms.date: 01/13/2020
ms.prod: non-product-specific
localization_priority: Priority
ms.openlocfilehash: 3c67fdb2b8582546c13624dcb8a6f139bb638df0
ms.sourcegitcommit: 0dacbe7c80ed387099e3ec21e151f8990b181ede
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/13/2020
ms.locfileid: "41111116"
---
# <a name="use-the-yeoman-generator-to-create-an-office-add-in-that-uses-single-sign-on-preview"></a><span data-ttu-id="cb9ff-103">Yeoman ジェネレーターを使用して、シングル サインオンを使用する Office アドインを作成する (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="cb9ff-103">Use the Yeoman generator to create an Office Add-in that uses single sign-on (preview)</span></span>

<span data-ttu-id="cb9ff-104">この記事では、可能な場合シングル サインオン (SSO) を使用し、SSO がサポートされていない場合は別のユーザー認証方法を使用する Excel、Word、または PowerPoint 用の Office アドインを作成するプロセスを説明します。 </span><span class="sxs-lookup"><span data-stu-id="cb9ff-104">In this article, you'll walk through the process of using the Yeoman generator to create an Office Add-in for Excel, Word, or PowerPoint that uses single sign-on (SSO) when possible, and uses an alternate method of user authentication when SSO is not supported.</span></span>

> [!TIP]
> <span data-ttu-id="cb9ff-105">このクイック スタートを完了する前に、「[Office アドインのシングル サインオンを有効化する](../develop/sso-in-office-add-ins.md)」を参照して、Office アドインの SSO に関する基本的な概念を確認してください。</span><span class="sxs-lookup"><span data-stu-id="cb9ff-105">Before you attempt to complete this quick start, review [Enable single sign-on for Office Add-ins](../develop/sso-in-office-add-ins.md) to learn basic concepts about SSO in Office Add-ins.</span></span> 
 
<span data-ttu-id="cb9ff-106">Yeoman ジェネレーターは、Azure 内で SSO を構成するために必要な手順を自動化し、SSO を使用するために必要なコードを生成することで、SSO アドインの作成プロセスを簡素化します。</span><span class="sxs-lookup"><span data-stu-id="cb9ff-106">The Yeoman generator simplifies the process of creating an SSO add-in, by automating the steps required to configure SSO within Azure and generating the code that's necessary for an add-in to use SSO.</span></span> <span data-ttu-id="cb9ff-107">Yeoman ジェネレーターが自動化する手順を手動で完了する方法についての詳細は、「[シングル サインオンを使用する Node.js Office アドインを作成する](../develop/create-sso-office-add-ins-nodejs.md)」チュートリアルを参照してください。</span><span class="sxs-lookup"><span data-stu-id="cb9ff-107">For a detailed walkthrough that describes how to manually complete the steps that the Yeoman generator automates, see the [Create a Node.js Office Add-in that uses single sign-on](../develop/create-sso-office-add-ins-nodejs.md) tutorial.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="cb9ff-108">前提条件</span><span class="sxs-lookup"><span data-stu-id="cb9ff-108">Prerequisites</span></span>

- <span data-ttu-id="cb9ff-109">[Node.js](https://nodejs.org) (バージョン 10.15.0 以降)</span><span class="sxs-lookup"><span data-stu-id="cb9ff-109">[Node.js](https://nodejs.org) (version 8.0.0 or later)</span></span>

- <span data-ttu-id="cb9ff-110">最新バージョンの [Yeoman](https://github.com/yeoman/yo) と [Office アドイン用の Yeoman ジェネレーター](https://github.com/OfficeDev/generator-office)。これらのツールをグローバルにインストールするには、コマンド プロンプトから次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="cb9ff-110">The latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office). To install these tools globally, run the following command via the command prompt:</span></span>

    ```command&nbsp;line
    npm install -g yo generator-office
    ```

    [!include[note to update Yeoman generator](../includes/note-yeoman-generator-update.md)]

- <span data-ttu-id="cb9ff-111">Office 365 (Office のサブスクリプション バージョン) アカウント。</span><span class="sxs-lookup"><span data-stu-id="cb9ff-111">Office 365 (the subscription version of Office).</span></span> <span data-ttu-id="cb9ff-112">Office 365 アカウントをまだお持ちでない場合は、[Office 365 開発者プログラム](https://aka.ms/devprogramsignup)に参加することで 90 日間の更新可能な無料の Office 365 サブスクリプションを入手できます。</span><span class="sxs-lookup"><span data-stu-id="cb9ff-112">If you don't already have an Office 365 account, you can get a free, 90-day renewable Office 365 subscription by joining the [Office 365 Developer Program](https://aka.ms/devprogramsignup).</span></span> 

- <span data-ttu-id="cb9ff-113">Insider の Office 365 のビルド。</span><span class="sxs-lookup"><span data-stu-id="cb9ff-113">An Insider's build of Office 365.</span></span> <span data-ttu-id="cb9ff-114">最新の月次バージョンを使用し、Insider チャネルからビルドする必要がありますが、このバージョンを入手するには [Office Insider である](https://products.office.com/office-insider?tab=tab-1)必要があります。</span><span class="sxs-lookup"><span data-stu-id="cb9ff-114">You should use the latest monthly version and build from the Insiders channel but you need to be an Office Insider to get this version.</span></span> 

    > [!NOTE]
    > <span data-ttu-id="cb9ff-115">ビルドが半期チャネルの運用に移行すると、そのビルドで SSO を含むプレビュー機能のサポートは無効になります。</span><span class="sxs-lookup"><span data-stu-id="cb9ff-115">Please note that when a build graduates to the production semi-annual channel, support for preview features, including SSO, is turned off for that build.</span></span>

## <a name="create-the-add-in-project"></a><span data-ttu-id="cb9ff-116">アドイン プロジェクトの作成</span><span class="sxs-lookup"><span data-stu-id="cb9ff-116">Create the add-in project</span></span>

> [!TIP]
> <span data-ttu-id="cb9ff-117">Yeoman ジェネレーターは、Excel、Word、または PowerPoint 用の SSO が有効な Office アドインを作成でき、JavaScript または TypeScript のスクリプト タイプで作成できます。</span><span class="sxs-lookup"><span data-stu-id="cb9ff-117">The Yeoman generator can create an SSO-enabled Office Add-in for Excel, Word, or PowerPoint, and can be created with script type of JavaScript or TypeScript.</span></span> <span data-ttu-id="cb9ff-118">次の手順では、`JavaScript` と `Excel` を指定しますが、使用しているシナリオに最適なスクリプト タイプと Office クライアント アプリケーションを選択する必要があります。</span><span class="sxs-lookup"><span data-stu-id="cb9ff-118">The following instructions specify `JavaScript` and `Excel`, but you should choose the script type and Office client application that best suits your scenario.</span></span>

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- <span data-ttu-id="cb9ff-119">**Choose a project type: (プロジェクトの種類を選択)** `Office Add-in Task Pane project supporting single sign-on`</span><span class="sxs-lookup"><span data-stu-id="cb9ff-119">**Choose a project type:** `Office Add-in Task Pane project supporting single sign-on`</span></span>
- <span data-ttu-id="cb9ff-120">**Choose a script type: (スクリプトの種類を選択)** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="cb9ff-120">**Choose a script type:** `Javascript`</span></span>
- <span data-ttu-id="cb9ff-121">**What would you want to name your add-in?: (アドインの名前を何にしますか)**</span><span class="sxs-lookup"><span data-stu-id="cb9ff-121">**What do you want to name your add-in?**</span></span> `My SSO Office Add-in`
- <span data-ttu-id="cb9ff-122">**Which Office client application would you like to support?: (どの Office クライアント アプリケーションをサポートしますか)**</span><span class="sxs-lookup"><span data-stu-id="cb9ff-122">**Which Office client application would you like to support?**</span></span> `Excel`

![Yeoman ジェネレーターのプロンプトと応答のスクリーンショット](../images/yo-office-sso-excel.png)

<span data-ttu-id="cb9ff-124">ウィザードを完了すると、ジェネレーターによってプロジェクトが作成されて、サポートしているノード コンポーネントがインストールされます。</span><span class="sxs-lookup"><span data-stu-id="cb9ff-124">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="explore-the-project"></a><span data-ttu-id="cb9ff-125">プロジェクトを確認する</span><span class="sxs-lookup"><span data-stu-id="cb9ff-125">Explore the project</span></span>

<span data-ttu-id="cb9ff-126">Yeoman ジェネレーターで作成したアドイン プロジェクトには、SSO が有効な作業ウィンドウ アドインのコードが含まれています。</span><span class="sxs-lookup"><span data-stu-id="cb9ff-126">The add-in project that you've created with the Yeoman generator contains sample code for a very basic task pane add-in.</span></span>

- <span data-ttu-id="cb9ff-127">プロジェクトのルートディレクトリにある **./ manifest.xml**ファイルは、アドインの設定と機能性を定義します。</span><span class="sxs-lookup"><span data-stu-id="cb9ff-127">The **./manifest.xml** file in the root directory of the project defines the settings and capabilities of the add-in.</span></span>

- <span data-ttu-id="cb9ff-128">**./src/taskpane/taskpane.html**ファイルには、作業ペイン用のHTMLマークアップが含まれています。</span><span class="sxs-lookup"><span data-stu-id="cb9ff-128">The **./src/taskpane/taskpane.html** file contains the HTML markup for the task pane.</span></span>
- <span data-ttu-id="cb9ff-129">**./src/taskpane/taskpane.css**ファイルには、作業ウィンドウ内のコンテンツに適用される CSS が含まれています。</span><span class="sxs-lookup"><span data-stu-id="cb9ff-129">The **./src/taskpane/taskpane.css** file contains the CSS that's applied to content in the task pane.</span></span>
- <span data-ttu-id="cb9ff-130">**./src/taskpane/taskpane.js**ファイルには、作業ウィンドウと Office のホスト アプリケーションの間のやり取りを容易にする Office JavaScript API コードが含まれています。</span><span class="sxs-lookup"><span data-stu-id="cb9ff-130">The **./src/taskpane/taskpane.js** file contains the Office JavaScript API code that facilitates interaction between the task pane and the Office host application.</span></span>

- <span data-ttu-id="cb9ff-131">**./src/helpers/documentHelper.js** ファイルは、Office JavaScript ライブラリを使用して、Microsoft Graph から Office ドキュメントにデータを追加します。</span><span class="sxs-lookup"><span data-stu-id="cb9ff-131">The **./src/helpers/documentHelper.js** file uses the Office JavaScript library to add the data from Microsoft Graph to the Office document.</span></span>
- <span data-ttu-id="cb9ff-132">**./src/helpers/fallbackauthdialog.html** ファイルは、フォールバック認証方法の JavaScript を読み込む UI を使用しないページです。</span><span class="sxs-lookup"><span data-stu-id="cb9ff-132">The **./src/helpers/fallbackauthdialog.html** file is the UI-less page that loads the fallback authentication method's JavaScript.</span></span>
- <span data-ttu-id="cb9ff-133">**./src/helpers/fallbackauthdialog.js** ファイルには、msal.js でユーザーにサインオンするフォールバック認証方法の JavaScript が含まれています。</span><span class="sxs-lookup"><span data-stu-id="cb9ff-133">The **./src/helpers/fallbackauthdialog.js** file contains the fallback authentication method's JavaScript that signs on the user with msal.js.</span></span>
- <span data-ttu-id="cb9ff-134">**./src/helpers/fallbackauthhelper.js** ファイルには、SSO 認証がサポートされていないシナリオでフォールバック認証方法を呼び出す作業ウィンドウの JavaScript が含まれています。</span><span class="sxs-lookup"><span data-stu-id="cb9ff-134">The **./src/helpers/fallbackauthhelper.js** file contains the task pane JavaScript that invokes the fallback authentication method in scenarios when SSO authentication is not supported.</span></span>
- <span data-ttu-id="cb9ff-135">**./src/helpers/ssoauthhelper.js** ファイルには、SSO API `getAccessToken` へのJavaScript 呼び出しが含まれ、ブートストラップ トークンの受信し、Microsoft Graph へのアクセス トークンのブートストラップ トークン交換の開始、データのための Microsoft Graph への呼び出しを行います。</span><span class="sxs-lookup"><span data-stu-id="cb9ff-135">The **./src/helpers/ssoauthhelper.js** file contains the JavaScript call to the SSO API, `getAccessToken`, receives the bootstrap token, initiates the swap of the bootstrap token for an access token to Microsoft Graph, and calls to Microsoft Graph for the data.</span></span>

- <span data-ttu-id="cb9ff-136">**./ENV** ファイルはプロジェクトのルート ディレクトリにあり、アドイン プロジェクトで使用される定数を定義します。</span><span class="sxs-lookup"><span data-stu-id="cb9ff-136">The **./ENV** file in the root directory of the project defines constants that are used by the add-in project.</span></span>
    > [!NOTE]
    > <span data-ttu-id="cb9ff-137">このファイルで定義されている定数の一部は、SSO プロセスを容易するために使用されます。</span><span class="sxs-lookup"><span data-stu-id="cb9ff-137">Some of the constants defined in this file are used to facilitate the SSO process.</span></span> <span data-ttu-id="cb9ff-138">このファイルの値を、特定のシナリオに合わせて更新できます。</span><span class="sxs-lookup"><span data-stu-id="cb9ff-138">You may want to update values in this file to match your specific scenario.</span></span> <span data-ttu-id="cb9ff-139">たとえば、アドインで `User.Read` 以外のものが必要な場合は、このファイルを更新して別の範囲を指定できます。</span><span class="sxs-lookup"><span data-stu-id="cb9ff-139">For example, you can update this file to specify a different scope, if your add-in requires something other than `User.Read`.</span></span>

## <a name="configure-sso"></a><span data-ttu-id="cb9ff-140">SSO を構成する</span><span class="sxs-lookup"><span data-stu-id="cb9ff-140">Configure SSO</span></span>

<span data-ttu-id="cb9ff-141">この時点では、アドイン プロジェクトが作成され、SSO プロセスを容易するために必要なコードが含まれています。</span><span class="sxs-lookup"><span data-stu-id="cb9ff-141">At this point, your add-in project has been created and contains the code that's necessary to facilitate the SSO process.</span></span> <span data-ttu-id="cb9ff-142">次の手順を完了して、アドインの SSO を構成します。</span><span class="sxs-lookup"><span data-stu-id="cb9ff-142">Next, complete the following steps to configure SSO for your add-in.</span></span>

1. <span data-ttu-id="cb9ff-143">プロジェクトのルート フォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="cb9ff-143">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "My SSO Office Add-in"
    ```

2. <span data-ttu-id="cb9ff-144">次のコマンドを実行して、アドインの SSO を構成します。</span><span class="sxs-lookup"><span data-stu-id="cb9ff-144">Run the following command to configure SSO for the add-in.</span></span>

    ```command&nbsp;line
    npm run configure-sso
    ```

    > [!WARNING]
    > <span data-ttu-id="cb9ff-145">このコマンドは、テナントが 2 要素認証を要求するように構成されている場合、失敗します。</span><span class="sxs-lookup"><span data-stu-id="cb9ff-145">This command will fail if your tenant is configured to require two-factor authentication.</span></span> <span data-ttu-id="cb9ff-146">このシナリオでは、「[シングル サインオンを使用する Node.js Office アドインを作成する](../develop/create-sso-office-add-ins-nodejs.md)」チュートリアルで説明されているように、Azure アプリの登録および SSO の構成手順を手動で完了する必要があります。</span><span class="sxs-lookup"><span data-stu-id="cb9ff-146">In this scenario, you'll need to manually complete the Azure app registration and SSO configuration steps, as described in the [Create a Node.js Office Add-in that uses single sign-on](../develop/create-sso-office-add-ins-nodejs.md) tutorial.</span></span>

3. <span data-ttu-id="cb9ff-147">Web ブラウザー ウィンドウが開き、Azure にサインインするように指示されます。</span><span class="sxs-lookup"><span data-stu-id="cb9ff-147">A web browser window will open and prompt you to sign in to Azure.</span></span> <span data-ttu-id="cb9ff-148">Office 365 管理者の資格情報を使用して Azure にサインインします。</span><span class="sxs-lookup"><span data-stu-id="cb9ff-148">Sign in to Union_Lite_2nd using your existing Office 365 administrator credentials.</span></span> <span data-ttu-id="cb9ff-149">これらの資格情報を使用して Azure に新しいアプリケーションが登録され、SSO に必要な設定が構成されます。</span><span class="sxs-lookup"><span data-stu-id="cb9ff-149">These credentials will be used to register a new application in Azure and configure the settings required by SSO.</span></span>

    > [!NOTE]
    > <span data-ttu-id="cb9ff-150">この手順で、管理者以外の資格情報を使用して Azure にサインインした場合、`configure-sso` スクリプトは、組織内のユーザーにアドインの管理者の同意を提供しません。</span><span class="sxs-lookup"><span data-stu-id="cb9ff-150">If you sign in to Azure using non-administrator credentials during this step, the `configure-sso` script won't be able to provide administrator consent for the add-in to users within your organization.</span></span> <span data-ttu-id="cb9ff-151">そのため、SSO はアドインのユーザーは使用できず、サインインするように求められます。</span><span class="sxs-lookup"><span data-stu-id="cb9ff-151">SSO will therefore not be available to users of the add-in and they'll be prompted to sign-in.</span></span>

4. <span data-ttu-id="cb9ff-152">資格情報を入力したら、ブラウザー ウィンドウを閉じ、コマンド プロンプトに戻ります。</span><span class="sxs-lookup"><span data-stu-id="cb9ff-152">After you enter your credentials, close the browser window and return to the command prompt.</span></span> <span data-ttu-id="cb9ff-153">SSO の構成プロセスが続行されると、コンソールに書き込まれたステータス メッセージが表示されます。</span><span class="sxs-lookup"><span data-stu-id="cb9ff-153">As the SSO configuration process continues, you'll see status messages being written to the console.</span></span> <span data-ttu-id="cb9ff-154">コンソール メッセージで説明されているように、Yeoman ジェネレーターが作成したアドイン プロジェクト内のファイルは、SSO プロセスで必要なデータで自動的に更新されます。</span><span class="sxs-lookup"><span data-stu-id="cb9ff-154">As described in the console messages, files within the add-in project that the Yeoman generator created are automatically updated with data that's required by the SSO process.</span></span>

## <a name="try-it-out"></a><span data-ttu-id="cb9ff-155">試してみる</span><span class="sxs-lookup"><span data-stu-id="cb9ff-155">Try it out</span></span>

1. <span data-ttu-id="cb9ff-156">SSO の構成プロセスが完了したら、次のコマンドを実行してプロジェクトを構築し、ローカル Web サーバーを起動して以前に選択した Office クライアント アプリケーションにアドインをサイドロードします。</span><span class="sxs-lookup"><span data-stu-id="cb9ff-156">When the SSO configuration process completes, run the following command to build the project, start the local web server, and sideload your add-in in the previously selected Office client application.</span></span>

    > [!NOTE]
    > <span data-ttu-id="cb9ff-157">開発の最中でも、OfficeアドインはHTTPではなくHTTPSを使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="cb9ff-157">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="cb9ff-158">次のコマンドを実行した後に証明書をインストールするように求められた場合は、Yeoman ジェネレーターによって提供される証明書をインストールするプロンプトを受け入れます。</span><span class="sxs-lookup"><span data-stu-id="cb9ff-158">If you are prompted to install a certificate after you run the following command, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

    ```command&nbsp;line
    npm start
    ```

2. <span data-ttu-id="cb9ff-159">前のコマンド (Excel、Word、PowerPoint など) を実行したときに開く Office クライアント アプリケーションで、[前のセクション](#configure-sso)の手順 3 で SSO を構成している間に Azure の接続に使用した Office 365 管理者アカウントと同じ Office 365 組織のメンバーであるユーザーでサインインしていることを確認します。</span><span class="sxs-lookup"><span data-stu-id="cb9ff-159">In the Office client application that opens when you run the previous command (i.e., Excel, Word or PowerPoint), make sure that you're signed in with a user that's a member of the same Office 365 organization as the Office 365 administrator account that you used to connect to Azure while configuring SSO in step 3 of the [previous section](#configure-sso).</span></span> <span data-ttu-id="cb9ff-160">これにより、SSO を正常に実行するための適切な条件が確立されます。</span><span class="sxs-lookup"><span data-stu-id="cb9ff-160">Doing so establishes the appropriate conditions for SSO to succeed.</span></span> 

3. <span data-ttu-id="cb9ff-161">Office クライアント アプリケーションで、[**ホーム**] タブを選択し、リボンの [**作業ウィンドウの表示**] ボタンをクリックして、アドインの作業ウィンドウを開きます。</span><span class="sxs-lookup"><span data-stu-id="cb9ff-161">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span> <span data-ttu-id="cb9ff-162">次の画像は、Excel のこのボタンを示しています。</span><span class="sxs-lookup"><span data-stu-id="cb9ff-162">The following image shows this button in Excel.</span></span>

    ![Excel アドイン ボタン](../images/excel-quickstart-addin-3b.png)

4. <span data-ttu-id="cb9ff-164">作業ウィンドウの下部にある [**マイ ユーザー プロファイルの情報を取得する**] ボタンを選択して、SSO プロセスを開始します。</span><span class="sxs-lookup"><span data-stu-id="cb9ff-164">At the bottom of the task pane, choose the **Get My User Profile Information** button to initiate the SSO process.</span></span> 

    > [!NOTE] 
    > <span data-ttu-id="cb9ff-165">この時点でまだ Office にサインインしていない場合は、サインインするように求められます。</span><span class="sxs-lookup"><span data-stu-id="cb9ff-165">If you're not already signed in to Office at this point, you'll be prompted to sign in.</span></span> <span data-ttu-id="cb9ff-166">前に説明したように、SSO を正常に実行するには、[前のセクション](#configure-sso)の手順 3 で SSO を構成している間に Azure の接続に使用した Office 365 管理者アカウントと同じ Office 365 組織のメンバーであるユーザーでサインインしている必要があります。</span><span class="sxs-lookup"><span data-stu-id="cb9ff-166">As described previously, you should sign in with a user that's a member of the same Office 365 organization as the Office 365 administrator account that you used to connect to Azure while configuring SSO in step 3 of the [previous section](#configure-sso), if you want SSO to succeed.</span></span>

5. <span data-ttu-id="cb9ff-167">アドインの代わりにアクセス許可を要求するダイアログ ウィンドウが表示される場合は、SSO はシナリオでサポートされず、代わりにアドインが別のユーザー認証方法に戻っていることを意味します。</span><span class="sxs-lookup"><span data-stu-id="cb9ff-167">If a dialog window appears to request permissions on behalf of the add-in, this means that SSO is not supported for your scenario and the add-in has instead fallen back to an alternate method of user authentication.</span></span> <span data-ttu-id="cb9ff-168">これは、テナント管理者がアドインが Microsoft Graph にアクセスするための同意を与えていない場合、または有効な Microsoft アカウントまたは Office 365 (「職場または学校」) アカウントで Office にサインインしていない場合に発生することがあります。</span><span class="sxs-lookup"><span data-stu-id="cb9ff-168">This may occur when the tenant administrator hasn't granted consent for the add-in to access Microsoft Graph, or when the user isn't signed into Office with a valid Microsoft Account or Office 365 ("Work or School") account.</span></span> <span data-ttu-id="cb9ff-169">ダイアログ ウィンドウで [**同意する**] ボタンを選択して続行します。</span><span class="sxs-lookup"><span data-stu-id="cb9ff-169">Choose the **Accept** button in the dialog window to continue.</span></span>

    ![アクセス許可を要求するダイアログ](../images/sso-permissions-request.png)

    > [!NOTE]
    > <span data-ttu-id="cb9ff-171">ユーザーがこのアクセス許可の要求を受け入れると、今後再びプロンプトが表示されることはありません。</span><span class="sxs-lookup"><span data-stu-id="cb9ff-171">After a user accepts this permissions request, they won't be prompted again in the future.</span></span>

6. <span data-ttu-id="cb9ff-172">アドインは、サインインしたユーザーのプロファイル情報を取得し、ドキュメントに書き込みます。</span><span class="sxs-lookup"><span data-stu-id="cb9ff-172">The add-in retrieves profile information for the signed-in user and writes it to the document.</span></span> <span data-ttu-id="cb9ff-173">次の画像は、Excel ワークシートに書き込まれたプロファイル情報の例を示します。</span><span class="sxs-lookup"><span data-stu-id="cb9ff-173">The following image shows an example of profile information written to an Excel worksheet.</span></span>

    ![Excel ワークシートのユーザー プロファイル情報](../images/sso-user-profile-info-excel.png)

## <a name="next-steps"></a><span data-ttu-id="cb9ff-175">次の手順</span><span class="sxs-lookup"><span data-stu-id="cb9ff-175">Next steps</span></span>

<span data-ttu-id="cb9ff-176">おめでとうございます。可能な場合 SSO を使用し、SSO がサポートされていない場合は別のユーザー認証方法を使用する作業ウィンドウ アドインを正常に作成しました。</span><span class="sxs-lookup"><span data-stu-id="cb9ff-176">Congratulations, you've successfully created a task pane add-in that uses SSO when possible, and uses an alternate method of user authentication when SSO is not supported.</span></span> <span data-ttu-id="cb9ff-177">Yeoman ジェネレーターが自動的に完了した SSO の構成手順、および SSO プロセスを容易にするコードの詳細については、「[シングル サインオンを使用する Node.js Office アドインを作成する](../develop/create-sso-office-add-ins-nodejs.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="cb9ff-177">To learn more about SSO configuration steps that the Yeoman generator completed automatically, and the code that facilitates the SSO process, see the [Create a Node.js Office Add-in that uses single sign-on](../develop/create-sso-office-add-ins-nodejs.md) tutorial.</span></span>

## <a name="see-also"></a><span data-ttu-id="cb9ff-178">関連項目</span><span class="sxs-lookup"><span data-stu-id="cb9ff-178">See also</span></span>

- [<span data-ttu-id="cb9ff-179">Office アドインのシングル サインオンを有効化する</span><span class="sxs-lookup"><span data-stu-id="cb9ff-179">Enable single sign-on for Office Add-ins</span></span>](../develop/sso-in-office-add-ins.md)
- [<span data-ttu-id="cb9ff-180">シングル サインオンを使用する Node.js Office アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="cb9ff-180">Create a Node.js Office Add-in that uses single sign-on</span></span>](../develop/create-sso-office-add-ins-nodejs.md)
- [<span data-ttu-id="cb9ff-181">シングル サインオン (SSO) のエラー メッセージのトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="cb9ff-181">Troubleshoot error messages for single sign-on (SSO)</span></span>](../develop/troubleshoot-sso-in-office-add-ins.md)