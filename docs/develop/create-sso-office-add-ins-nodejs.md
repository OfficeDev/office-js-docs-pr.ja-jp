---
title: シングル サインオンを使用する Node.js Office アドインを作成する
description: Office シングル サインオンを使用する Node.js ベースのアドインを作成する方法を学ぶ
ms.date: 07/30/2020
localization_priority: Normal
ms.openlocfilehash: a587f6a1ad596f4006734e0d850e009c3fc145b0
ms.sourcegitcommit: 6e6de48f746416ec68b2cf4c298253986486fbfd
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/14/2020
ms.locfileid: "47651938"
---
# <a name="create-a-nodejs-office-add-in-that-uses-single-sign-on"></a><span data-ttu-id="a3291-103">シングル サインオンを使用する Node.js Office アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="a3291-103">Create a Node.js Office Add-in that uses single sign-on</span></span>

<span data-ttu-id="a3291-p101">ユーザーは、このサインイン プロセスを利用してユーザーを承認する Office および Office Web アドインにサインインできます。こうして承認されたユーザーは、アドインと Microsoft Graph への 2 度目のサインオンの必要がなくなります。概要については、「[Office アドインで SSO を有効化する](sso-in-office-add-ins.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="a3291-p101">Users can sign in to Office, and your Office Web Add-in can take advantage of this sign-in process to authorize users to your add-in and to Microsoft Graph without requiring users to sign in a second time. For an overview, see [Enable SSO in an Office Add-in](sso-in-office-add-ins.md).</span></span>

<span data-ttu-id="a3291-106">この記事では、Node.js と Express を使用して作成したアドインで、シングル サインオン (SSO) を有効化するプロセスについて手順を追って説明します。</span><span class="sxs-lookup"><span data-stu-id="a3291-106">This article walks you through the process of enabling single sign-on (SSO) in an add-in that is built with Node.js and Express.</span></span> <span data-ttu-id="a3291-107">ASP.NET ベースのアドインに関する同様の記事については、「[シングル サインオンを使用する ASP.NET Office アドインを作成する](create-sso-office-add-ins-aspnet.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="a3291-107">For a similar article about an ASP.NET-based add-in, see [Create an ASP.NET Office Add-in that uses single sign-on](create-sso-office-add-ins-aspnet.md).</span></span>

> [!NOTE]
> <span data-ttu-id="a3291-108">この記事で説明する手順を完了する代わりに、Yeoman ジェネレーターを使用して SSO が有効な Node.js Office アドインを作成することもできます。</span><span class="sxs-lookup"><span data-stu-id="a3291-108">As an alternative to completing the steps described in this article, you can use the Yeoman generator to create an SSO-enabled, Node.js Office Add-in.</span></span> <span data-ttu-id="a3291-109">Yeoman ジェネレーターは、Azure 内で SSO を構成するために必要な手順を自動化し、SSO を使用するために必要なコードを生成することで、SSO が有効なアドインの作成プロセスを簡素化します。</span><span class="sxs-lookup"><span data-stu-id="a3291-109">The Yeoman generator simplifies the process of creating an SSO-enabled add-in, by automating the steps required to configure SSO within Azure and generating the code that's necessary for an add-in to use SSO.</span></span> <span data-ttu-id="a3291-110">詳細については、「[シングル サインオン (SSO) のクイック スタート](../quickstarts/sso-quickstart.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="a3291-110">For more information, see the [Single sign-on (SSO) quick start](../quickstarts/sso-quickstart.md).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="a3291-111">前提条件</span><span class="sxs-lookup"><span data-stu-id="a3291-111">Prerequisites</span></span>

* <span data-ttu-id="a3291-112">[Node.js](https://nodejs.org/) (最新 [LTS](https://nodejs.org/about/releases) バージョン)</span><span class="sxs-lookup"><span data-stu-id="a3291-112">[Node.js](https://nodejs.org/) (the latest [LTS](https://nodejs.org/about/releases) version)</span></span>

* <span data-ttu-id="a3291-113">[Git バッシュ](https://git-scm.com/downloads) (またはその他の Git クライアント)</span><span class="sxs-lookup"><span data-stu-id="a3291-113">[Git Bash](https://git-scm.com/downloads) (or another git client)</span></span>

* <span data-ttu-id="a3291-114">TypeScript、バージョン 3.6.2 以降</span><span class="sxs-lookup"><span data-stu-id="a3291-114">TypeScript, version 3.6.2 or later</span></span>

[!include[additional prerequisites](../includes/sso-tutorial-prereqs.md)]

* <span data-ttu-id="a3291-115">コード エディター。</span><span class="sxs-lookup"><span data-stu-id="a3291-115">A code editor.</span></span> <span data-ttu-id="a3291-116">Visual Studio Code をお勧めします。</span><span class="sxs-lookup"><span data-stu-id="a3291-116">We recommend Visual Studio Code.</span></span>

* <span data-ttu-id="a3291-117">少なくとも、Microsoft 365 サブスクリプションの OneDrive for Business に格納されているファイルとフォルダーがいくつかあります。</span><span class="sxs-lookup"><span data-stu-id="a3291-117">At least a few files and folders stored on OneDrive for Business in your Microsoft 365 subscription.</span></span>

* <span data-ttu-id="a3291-118">Microsoft Azure サブスクリプション。</span><span class="sxs-lookup"><span data-stu-id="a3291-118">A Microsoft Azure subscription.</span></span> <span data-ttu-id="a3291-119">このアドインには、Azure Active Directory (AD) が必要です。</span><span class="sxs-lookup"><span data-stu-id="a3291-119">This add-in requires Azure Active Directory (AD).</span></span> <span data-ttu-id="a3291-120">Azure AD は、アプリケーションが認証および承認に使用する ID サービスを提供します。</span><span class="sxs-lookup"><span data-stu-id="a3291-120">Azure AD provides identity services that applications use for authentication and authorization.</span></span> <span data-ttu-id="a3291-121">[Microsoft Azure](https://account.windowsazure.com/SignUp) で試用版サブスクリプションを取得できます。</span><span class="sxs-lookup"><span data-stu-id="a3291-121">A trial subscription can be acquired at [Microsoft Azure](https://account.windowsazure.com/SignUp).</span></span>

## <a name="set-up-the-starter-project"></a><span data-ttu-id="a3291-122">スタート プロジェクトをセットアップする</span><span class="sxs-lookup"><span data-stu-id="a3291-122">Set up the starter project</span></span>

1. <span data-ttu-id="a3291-123">「[Office Add-in NodeJS SSO](https://github.com/officedev/office-add-in-nodejs-sso)」にあるリポジトリを複製するかダウンロードします。</span><span class="sxs-lookup"><span data-stu-id="a3291-123">Clone or download the repo at [Office Add-in NodeJS SSO](https://github.com/officedev/office-add-in-nodejs-sso).</span></span>

    > [!NOTE]
    > <span data-ttu-id="a3291-124">このサンプルには、次の 3 つのバージョンがあります。</span><span class="sxs-lookup"><span data-stu-id="a3291-124">There are three versions of the sample:</span></span>  
    > * <span data-ttu-id="a3291-p106">**開始**フォルダーはスタートプロジェクトです。SSO または承認に直接接続されていないアドインの UI とその他の側面は、既に実行されています。この記事の後のセクションでは、これを完了するプロセスについて順を追って説明します。</span><span class="sxs-lookup"><span data-stu-id="a3291-p106">The **Begin** folder is a starter project. The UI and other aspects of the add-in that are not directly connected to SSO or authorization are already done. Later sections of this article walk you through the process of completing it.</span></span>
    > * <span data-ttu-id="a3291-128">このサンプルの **[Complete]** バージョンは、この記事の手順を完了したときに得られるアドインと同様のものですが、完成済みのプロジェクトには、この記事のテキストと重複するコード コメントが含まれています。</span><span class="sxs-lookup"><span data-stu-id="a3291-128">The **Complete** version of the sample is just like the add-in that you would have if you completed the procedures of this article, except that the completed project has code comments that would be redundant with the text of this article.</span></span> <span data-ttu-id="a3291-129">完成したバージョンを使用するには、この記事に記載されている手順に従いますが、"Begin" を "Completed" に置き換え、 **クライアント側のコードを記述** してサーバー側を **コーディング** するセクションをスキップします。</span><span class="sxs-lookup"><span data-stu-id="a3291-129">To use the completed version, just follow the instructions in this article, but replace "Begin" with "Completed" and skip the sections **Code the client side** and **Code the server** side.</span></span>
    > * <span data-ttu-id="a3291-130">**SSOAutoSetup** バージョンは、アドインを Azure AD に登録して構成する手順の大部分を自動化する完成されたサンプルです。</span><span class="sxs-lookup"><span data-stu-id="a3291-130">The **SSOAutoSetup** version is a completed sample that automates most of the steps to register the add-in with Azure AD and configure it.</span></span> <span data-ttu-id="a3291-131">SSO で動作するアドインをすばやく表示する場合には、このバージョンを使用します。</span><span class="sxs-lookup"><span data-stu-id="a3291-131">Use this version if you want to see a working add-in with SSO quickly.</span></span> <span data-ttu-id="a3291-132">フォルダーの Readme の手順に従ってください。</span><span class="sxs-lookup"><span data-stu-id="a3291-132">Just follow the steps in the Readme of the folder.</span></span> <span data-ttu-id="a3291-133">Azure AD とアドインの関係をよりよく理解するために、この記事にある手動での登録およびセットアップのステップを行うことをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="a3291-133">We recommend that at some point you go through the manual registration and setup steps in this article to better understand the relationship between Azure AD and an add-in.</span></span> 

1. <span data-ttu-id="a3291-134">**開始**フォルダーでコマンドプロンプトを開きます。</span><span class="sxs-lookup"><span data-stu-id="a3291-134">Open a command prompt in the **Begin** folder.</span></span>

1. <span data-ttu-id="a3291-135">コンソールで `npm install` を入力して、package.json ファイルに項目化されているすべての依存関係をインストールします。</span><span class="sxs-lookup"><span data-stu-id="a3291-135">Enter `npm install` in the console to install all of the dependencies itemized in the package.json file.</span></span>

1. <span data-ttu-id="a3291-136">コマンド`npm run install-dev-certs`を実行します。</span><span class="sxs-lookup"><span data-stu-id="a3291-136">Run the command `npm run install-dev-certs`.</span></span> <span data-ttu-id="a3291-137">証明書をインストールするプロンプトに対して**はい**を選択します。</span><span class="sxs-lookup"><span data-stu-id="a3291-137">Select **Yes** to the prompt to install the certificate.</span></span>

## <a name="register-the-add-in-with-azure-ad-v20-endpoint"></a><span data-ttu-id="a3291-138">Azure AD v2.0 エンドポイントにアドインを登録する</span><span class="sxs-lookup"><span data-stu-id="a3291-138">Register the add-in with Azure AD v2.0 endpoint</span></span>

1. <span data-ttu-id="a3291-139">[Azure ポータル - アプリの登録](https://go.microsoft.com/fwlink/?linkid=2083908)ページに移動してアプリを登録します。</span><span class="sxs-lookup"><span data-stu-id="a3291-139">Navigate to the [Azure portal - App registrations](https://go.microsoft.com/fwlink/?linkid=2083908) page to register your app.</span></span>

1. <span data-ttu-id="a3291-140">Microsoft 365 テナントに対して ***管理者*** の資格情報を使用してサインインします。</span><span class="sxs-lookup"><span data-stu-id="a3291-140">Sign in with the ***admin*** credentials to your Microsoft 365 tenancy.</span></span> <span data-ttu-id="a3291-141">たとえば、MyName@contoso.onmicrosoft.com です。</span><span class="sxs-lookup"><span data-stu-id="a3291-141">For example, MyName@contoso.onmicrosoft.com.</span></span>

1. <span data-ttu-id="a3291-142">**[新規登録]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="a3291-142">Select **New registration**.</span></span> <span data-ttu-id="a3291-143">**[アプリケーションを登録]** ページで、次のように値を設定します。</span><span class="sxs-lookup"><span data-stu-id="a3291-143">On the **Register an application** page, set the values as follows.</span></span>

    * <span data-ttu-id="a3291-144">`Office-Add-in-NodeJS-SSO` に **[名前]** を設定します。</span><span class="sxs-lookup"><span data-stu-id="a3291-144">Set **Name** to `Office-Add-in-NodeJS-SSO`.</span></span>
    * <span data-ttu-id="a3291-145">**[サポートされているアカウントの種類]** を **[任意の組織のディレクトリ内のアカウントと個人用の Microsoft アカウント (例: Skype、 Xbox、Outlook.com)]** に設定します。</span><span class="sxs-lookup"><span data-stu-id="a3291-145">Set **Supported account types** to **Accounts in any organizational directory and personal Microsoft accounts (e.g. Skype, Xbox, Outlook.com)**.</span></span>
    * <span data-ttu-id="a3291-146">アプリケーションの種類を [ **Web** ] に設定し、 **リダイレクト URI** をに設定し ` https://localhost:44355/dialog.html` ます。</span><span class="sxs-lookup"><span data-stu-id="a3291-146">Set the application type to **Web** and then set **Redirect URI** to ` https://localhost:44355/dialog.html`.</span></span>
    * <span data-ttu-id="a3291-147">**[登録]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="a3291-147">Choose **Register**.</span></span>

1. <span data-ttu-id="a3291-148">**Office-Add-in-NodeJS-SSO** ページで、**アプリケーション (クライアント) ID** と**ディレクトリ (テナント) ID** の値をコピーして保存します。</span><span class="sxs-lookup"><span data-stu-id="a3291-148">On the **Office-Add-in-NodeJS-SSO** page, copy and save the values for the **Application (client) ID** and the **Directory (tenant) ID**.</span></span> <span data-ttu-id="a3291-149">以降の手順では、それらの両方を使用します。</span><span class="sxs-lookup"><span data-stu-id="a3291-149">You'll use both of them in later procedures.</span></span>

    > [!NOTE]
    > <span data-ttu-id="a3291-150">この **アプリケーション (クライアント) ID** は、Office クライアントアプリケーション (PowerPoint、Word、Excel など) などの他のアプリケーションがアプリケーションへの承認されたアクセスをシークするときの "audience" 値です。</span><span class="sxs-lookup"><span data-stu-id="a3291-150">This **Application (client) ID** is the "audience" value when other applications, such as the Office client application (e.g., PowerPoint, Word, Excel), seek authorized access to the application.</span></span> <span data-ttu-id="a3291-151">また、そのアプリケーションが Microsoft Graph への承認されたアクセスを求めるときには、このアプリケーションの「クライアント ID」になります。</span><span class="sxs-lookup"><span data-stu-id="a3291-151">It is also the "client ID" of the application when it, in turn, seeks authorized access to Microsoft Graph.</span></span>

1. <span data-ttu-id="a3291-152">**[管理]** の下の **[認証]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="a3291-152">Select **Authentication** under **Manage**.</span></span> <span data-ttu-id="a3291-153">[ **暗黙的な付与** ] セクションで、 **アクセストークン** と **ID トークン**の両方のチェックボックスをオンにします。</span><span class="sxs-lookup"><span data-stu-id="a3291-153">In the **Implicit grant** section, enable the checkboxes for both **Access token** and **ID token**.</span></span> <span data-ttu-id="a3291-154">サンプルには、SSO が利用できないときに呼び出されるフォールバック認証システムがあります。</span><span class="sxs-lookup"><span data-stu-id="a3291-154">The sample has a fallback authorization system that is invoked when SSO is not available.</span></span> <span data-ttu-id="a3291-155">このシステムは、暗黙的フローを使用します。</span><span class="sxs-lookup"><span data-stu-id="a3291-155">This system uses the Implicit Flow.</span></span>

1. <span data-ttu-id="a3291-156">フォームの最上部で **[保存]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="a3291-156">Select **Save** at the top of the form.</span></span>

1. <span data-ttu-id="a3291-157">**[管理]** で **[証明書とシークレット]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="a3291-157">Select **Certificates & secrets** under **Manage**.</span></span> <span data-ttu-id="a3291-158">**[新しいクライアント シークレット]** ボタンを選択します。</span><span class="sxs-lookup"><span data-stu-id="a3291-158">Select the **New client secret** button.</span></span> <span data-ttu-id="a3291-159">**[説明]** に値を入力してから、**[有効期限]** に適切なオプションを選択し、**[追加]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="a3291-159">Enter a value for **Description** then select an appropriate option for **Expires** and choose **Add**.</span></span> <span data-ttu-id="a3291-160">*クライアント シークレットの値をすぐにコピーして、後の手順で必要になるため、先に進む前にアプリケーションIDと一緒に保存*してください。</span><span class="sxs-lookup"><span data-stu-id="a3291-160">*Copy the client secret value immediately and save it with the application ID* before proceeding as you'll need it in a later procedure.</span></span>

1. <span data-ttu-id="a3291-161">**[管理]** の下の **[API の公開]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="a3291-161">Select **Expose an API** under **Manage**.</span></span> <span data-ttu-id="a3291-162">[ **設定** ] リンクを選択します。</span><span class="sxs-lookup"><span data-stu-id="a3291-162">Select the **Set** link.</span></span> <span data-ttu-id="a3291-163">これにより、"api://$App ID GUID $" という形式のアプリケーション ID URI が生成されます。ここで、$App ID GUID $ は **アプリケーション (クライアント) ID**です。</span><span class="sxs-lookup"><span data-stu-id="a3291-163">This will generate the Application ID URI in the form "api://$App ID GUID$", where $App ID GUID$ is the **Application (client) ID**.</span></span>

1. <span data-ttu-id="a3291-164">生成された ID で、を挿入し `localhost:44355/` ます (末尾にスラッシュ "/" を追加します)。二重スラッシュと GUID の間に追加します。</span><span class="sxs-lookup"><span data-stu-id="a3291-164">In the generated ID, insert `localhost:44355/` (note the forward slash "/" appended to the end) between the double forward slashes and the GUID.</span></span> <span data-ttu-id="a3291-165">完了すると、ID 全体にフォームが表示され `api://localhost:44355/$App ID GUID$` ます。たとえば、次のようになり `api://localhost:44355/c6c1f32b-5e55-4997-881a-753cc1d563b7` ます。</span><span class="sxs-lookup"><span data-stu-id="a3291-165">When you are finished, the entire ID should have the form `api://localhost:44355/$App ID GUID$`; for example `api://localhost:44355/c6c1f32b-5e55-4997-881a-753cc1d563b7`.</span></span>

1. <span data-ttu-id="a3291-166">**[Scope の追加]** ボタンをクリックします。</span><span class="sxs-lookup"><span data-stu-id="a3291-166">Select the **Add a scope** button.</span></span> <span data-ttu-id="a3291-167">開いたパネルで、`access_as_user`を **[スコープ名]** として入力します。</span><span class="sxs-lookup"><span data-stu-id="a3291-167">In the panel that opens, enter `access_as_user` as the **Scope** name.</span></span>

1. <span data-ttu-id="a3291-168">**[同意できるのはだれですか?]** を **[管理者とユーザー]** に設定します。</span><span class="sxs-lookup"><span data-stu-id="a3291-168">Set **Who can consent?** to **Admins and users**.</span></span>

1. <span data-ttu-id="a3291-169">ユーザーが `access_as_user` 現在のユーザーと同じ権限で Office クライアントアプリケーションでアドインの Web api を使用できるようにするには、[管理者] と [ユーザーの同意] を構成するためのフィールドに必要な値を入力します。</span><span class="sxs-lookup"><span data-stu-id="a3291-169">Fill in the fields for configuring the admin and user consent prompts with values that are appropriate for the `access_as_user` scope which enables the Office client application to use your add-in's web APIs with the same rights as the current user.</span></span> <span data-ttu-id="a3291-170">提案:</span><span class="sxs-lookup"><span data-stu-id="a3291-170">Suggestions:</span></span>

    - <span data-ttu-id="a3291-171">**管理者の同意の表示名**: Office はユーザーとして機能します。</span><span class="sxs-lookup"><span data-stu-id="a3291-171">**Admin consent display name**: Office can act as the user.</span></span>
    - <span data-ttu-id="a3291-172">**管理者の同意の説明**: 現在のユーザーと同じ権限で Office がアドインの Web API を呼び出すことを可能にします。</span><span class="sxs-lookup"><span data-stu-id="a3291-172">**Admin consent description**: Enable Office to call the add-in's web APIs with the same rights as the current user.</span></span>
    - <span data-ttu-id="a3291-173">**ユーザーの同意の表示名**: Office はあなたとして機能します。</span><span class="sxs-lookup"><span data-stu-id="a3291-173">**User consent display name**: Office can act as you.</span></span>
    - <span data-ttu-id="a3291-174">**ユーザーの同意の説明**: 自分と同じ権限でアドインの web api を呼び出すように Office を有効にします。</span><span class="sxs-lookup"><span data-stu-id="a3291-174">**User consent description**: Enable Office to call the add-in's web APIs with the same rights that you have.</span></span>

1. <span data-ttu-id="a3291-175">**[状態]** が **[有効]** に設定されていることを確認してください。</span><span class="sxs-lookup"><span data-stu-id="a3291-175">Ensure that **State** is set to **Enabled**.</span></span>

1. <span data-ttu-id="a3291-176">**[スコープの追加]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="a3291-176">Select **Add scope** .</span></span>

    > [!NOTE]
    > <span data-ttu-id="a3291-177">テキストフィールドのすぐ下に表示される **[スコープ名]** のドメイン部分は、以前に設定したアプリケーション ID URI に自動的に一致し、末尾に`/access_as_user`が追加されます。たとえば、`api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`です。</span><span class="sxs-lookup"><span data-stu-id="a3291-177">The domain part of the **Scope** name displayed just below the text field should automatically match the Application ID URI that you set earlier, with `/access_as_user` appended to the end; for example, `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`.</span></span>

1. <span data-ttu-id="a3291-178">**[承認済みのクライアント アプリケーション]** セクションで、アドインの Web アプリケーションに対して承認するアプリケーションを特定します。</span><span class="sxs-lookup"><span data-stu-id="a3291-178">In the **Authorized client applications** section, you identify the applications that you want to authorize to your add-in's web application.</span></span> <span data-ttu-id="a3291-179">次のそれぞれの ID を事前承認する必要があります。</span><span class="sxs-lookup"><span data-stu-id="a3291-179">Each of the following IDs needs to be pre-authorized.</span></span>

    - <span data-ttu-id="a3291-180">`d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)</span><span class="sxs-lookup"><span data-stu-id="a3291-180">`d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)</span></span>
    - <span data-ttu-id="a3291-181">`ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (Microsoft Office)</span><span class="sxs-lookup"><span data-stu-id="a3291-181">`ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (Microsoft Office)</span></span>
    - <span data-ttu-id="a3291-182">`57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office on the web)</span><span class="sxs-lookup"><span data-stu-id="a3291-182">`57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office on the web)</span></span>
    - <span data-ttu-id="a3291-183">`08e18876-6177-487e-b8b5-cf950c1e598c` (Office on the web)</span><span class="sxs-lookup"><span data-stu-id="a3291-183">`08e18876-6177-487e-b8b5-cf950c1e598c` (Office on the web)</span></span>
    - <span data-ttu-id="a3291-184">`bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Outlook on the web)</span><span class="sxs-lookup"><span data-stu-id="a3291-184">`bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Outlook on the web)</span></span>

    <span data-ttu-id="a3291-185">ID ごとに、次の手順を実行します。</span><span class="sxs-lookup"><span data-stu-id="a3291-185">For each ID, take these steps:</span></span>

    <span data-ttu-id="a3291-186">a. </span><span class="sxs-lookup"><span data-stu-id="a3291-186">a.</span></span> <span data-ttu-id="a3291-187">**[クライアント アプリケーションの追加]** ボタンを選択し、表示されたパネルで [クライアント ID] をそれぞれの GUID に設定して、`api://localhost:44355/$App ID GUID$/access_as_user`のチェック ボックスをオンにします。</span><span class="sxs-lookup"><span data-stu-id="a3291-187">Select **Add a client application** button and then, in the panel that opens, set the Client ID to the respective GUID and check the box for `api://localhost:44355/$App ID GUID$/access_as_user`.</span></span>

    <span data-ttu-id="a3291-188">b. </span><span class="sxs-lookup"><span data-stu-id="a3291-188">b.</span></span> <span data-ttu-id="a3291-189">**[アプリケーションの追加]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="a3291-189">Select **Add application**.</span></span>

1. <span data-ttu-id="a3291-190">**[管理]** の下の **[API アクセス許可]** を選択し、**[アクセス許可の追加]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="a3291-190">Select **API permissions** under **Manage** and select **Add a permission**.</span></span> <span data-ttu-id="a3291-191">開いたパネルで、**[Microsoft Graph]** を選択してから **[委任されたアクセス許可]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="a3291-191">On the panel that opens, choose **Microsoft Graph** and then choose **Delegated permissions**.</span></span>

1. <span data-ttu-id="a3291-192">アドインに必要な権限を検索するには、**[アクセス許可を選択]** の検索ボックスを使用します。</span><span class="sxs-lookup"><span data-stu-id="a3291-192">Use the **Select permissions** search box to search for the permissions your add-in needs.</span></span> <span data-ttu-id="a3291-193">以下を選択します。</span><span class="sxs-lookup"><span data-stu-id="a3291-193">Select the following.</span></span> <span data-ttu-id="a3291-194">アドイン自体に必要なのは最初のものだけです。ただし、 `profile` Office アプリケーションがアドイン web アプリケーションへのトークンを取得するには、そのアクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="a3291-194">Only the first is really required by your add-in itself; but the `profile` permission is required for the Office application to get a token to your add-in web application.</span></span>

    * <span data-ttu-id="a3291-195">Files.Read.All</span><span class="sxs-lookup"><span data-stu-id="a3291-195">Files.Read.All</span></span>
    * <span data-ttu-id="a3291-196">profile</span><span class="sxs-lookup"><span data-stu-id="a3291-196">profile</span></span>

    > [!NOTE]
    > <span data-ttu-id="a3291-197">`User.Read` アクセス許可は既定でリストされています。</span><span class="sxs-lookup"><span data-stu-id="a3291-197">The `User.Read` permission may already be listed by default.</span></span> <span data-ttu-id="a3291-198">必要でないアクセス許可は依頼しない方がよいため、アドインが実際に必要でない場合は、このアクセス許可のボックスのチェックをオフにしておくことをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="a3291-198">It is a good practice not to ask for permissions that are not needed, so we recommend that you uncheck the box for this permission if your add-in does not actually need it.</span></span>

1. <span data-ttu-id="a3291-199">表示される各アクセス許可のチェック ボックスをオンにします。</span><span class="sxs-lookup"><span data-stu-id="a3291-199">Select the check box for each permission as it appears.</span></span> <span data-ttu-id="a3291-200">アドインに必要なアクセス許可を選択したら、パネルの下部にある **[アクセス許可を追加する]** ボタンをクリックします。</span><span class="sxs-lookup"><span data-stu-id="a3291-200">After selecting the permissions that your add-in needs, select the **Add permissions** button at the bottom of the panel.</span></span>

1. <span data-ttu-id="a3291-201">同じページで、**[[テナント名]に管理者の同意を与える]** ボタンを選択し、表示される確認に対して **[はい]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="a3291-201">On the same page, choose the **Grant admin consent for [tenant name]** button, and then select **Yes** for the confirmation that appears.</span></span>

## <a name="configure-the-add-in"></a><span data-ttu-id="a3291-202">アドインを構成する</span><span class="sxs-lookup"><span data-stu-id="a3291-202">Configure the add-in</span></span>

1. <span data-ttu-id="a3291-203">コード エディターで複製プロジェクトの`\Begin`フォルダーを開きます。</span><span class="sxs-lookup"><span data-stu-id="a3291-203">Open the `\Begin` folder in the cloned project in your code editor.</span></span>

1. <span data-ttu-id="a3291-204">`.ENV`ファイルを開き、以前にコピーした値を使用します。</span><span class="sxs-lookup"><span data-stu-id="a3291-204">Open the `.ENV` file and use the values that you copied earlier.</span></span> <span data-ttu-id="a3291-205">**CLIENT_ID** を**アプリケーション (クライアント) ID** に設定し、**CLIENT_SECRET** をクライアント シークレットに設定します。</span><span class="sxs-lookup"><span data-stu-id="a3291-205">Set the **CLIENT_ID** to your **Application (client) ID**, and set the **CLIENT_SECRET** to your client secret.</span></span> <span data-ttu-id="a3291-206">値は引用符で囲ま**ない**でください。</span><span class="sxs-lookup"><span data-stu-id="a3291-206">The values should **not** be in quotation marks.</span></span> <span data-ttu-id="a3291-207">完了すると、ファイルは以下のようになります。</span><span class="sxs-lookup"><span data-stu-id="a3291-207">When you are done, the file should be similar to the following:</span></span> 

    ```javascript
    CLIENT_ID=8791c036-c035-45eb-8b0b-265f43cc4824
    CLIENT_SECRET=X7szTuPwKNts41:-/fa3p.p@l6zsyI/p
    NODE_ENV=development
    ```

1. <span data-ttu-id="a3291-208">`\public\javascripts\fallbackAuthDialog.js`ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="a3291-208">Open the `\public\javascripts\fallbackAuthDialog.js` file.</span></span> <span data-ttu-id="a3291-209">`msalConfig`宣言では、プレースホルダー $application_GUID here$ はアドインの登録時にコピーしたアプリケーション ID に置き換えます。</span><span class="sxs-lookup"><span data-stu-id="a3291-209">In the `msalConfig` declaration, replace the placeholder $application_GUID here$ with the Application ID that you copied when you registered your add-in.</span></span> <span data-ttu-id="a3291-210">値は引用符で囲む必要があります。</span><span class="sxs-lookup"><span data-stu-id="a3291-210">The value should be in quotation marks.</span></span>

1. <span data-ttu-id="a3291-211">アドイン マニフェスト ファイル "manifest\manifest_local.xml" を開き、ファイルの一番下までスクロールします。</span><span class="sxs-lookup"><span data-stu-id="a3291-211">Open the add-in manifest file "manifest\manifest_local.xml" and then scroll to the bottom of the file.</span></span> <span data-ttu-id="a3291-212">`</VersionOverrides>`終了タグのすぐ上に、以下のマークアップがあります。</span><span class="sxs-lookup"><span data-stu-id="a3291-212">Just above the `</VersionOverrides>` end tag, you'll find the following markup:</span></span>

    ```xml
    <WebApplicationInfo>
      <Id>$application_GUID here$</Id>
      <Resource>api://localhost:44355/$application_GUID here$</Resource>
      <Scopes>
          <Scope>Files.Read.All</Scope>
          <Scope>profile</Scope>
      </Scopes>
    </WebApplicationInfo>
    ```

1. <span data-ttu-id="a3291-213">このマークアップ内の*両方の場所の*プレースホルダー "$application_GUID here$" を、アドインの登録時にコピーしたアプリケーション ID に置き換えます。</span><span class="sxs-lookup"><span data-stu-id="a3291-213">Replace the placeholder "$application_GUID here$" *in both places* in the markup with the Application ID that you copied when you registered your add-in.</span></span> <span data-ttu-id="a3291-214">"$" 記号は ID の一部ではないため、含めないでください。</span><span class="sxs-lookup"><span data-stu-id="a3291-214">The "$" symbols are not part of the ID, so do not include them.</span></span> <span data-ttu-id="a3291-215">これは、の CLIENT_ID と対象ユーザーに対して使用したものと同じ ID です。ENV ファイル。</span><span class="sxs-lookup"><span data-stu-id="a3291-215">This is the same ID you used in for the CLIENT_ID and Audience in the .ENV file.</span></span>

    > [!NOTE]
    > <span data-ttu-id="a3291-216">**リソース**値は、アドインを登録したときに設定した**アプリケーション ID URI** です。</span><span class="sxs-lookup"><span data-stu-id="a3291-216">The **Resource** value is the **Application ID URI** you set when you registered the add-in.</span></span> <span data-ttu-id="a3291-217">**[範囲]** セクションは、アドインが AppSource を通じて販売される場合に同意ダイアログ ボックスを生成するためにのみ使用されます。</span><span class="sxs-lookup"><span data-stu-id="a3291-217">The **Scopes** section is used only to generate a consent dialog box if the add-in is sold through AppSource.</span></span>

## <a name="code-the-client-side"></a><span data-ttu-id="a3291-218">クライアント側のコーディング</span><span class="sxs-lookup"><span data-stu-id="a3291-218">Code the client-side</span></span>

### <a name="create-the-sso-logic"></a><span data-ttu-id="a3291-219">SSO ロジックを作成する</span><span class="sxs-lookup"><span data-stu-id="a3291-219">Create the SSO logic</span></span>

1. <span data-ttu-id="a3291-220">コード エディターで、`public\javascripts\ssoAuthES6.js`ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="a3291-220">In your code editor, open the file `public\javascripts\ssoAuthES6.js`.</span></span> <span data-ttu-id="a3291-221">Internet Explorer 11 でも Promise がサポートされることを保証するコードと、アドインの唯一のボタンにハンドラーを割り当てるための`Office.onReady`呼び出しが既にあります。</span><span class="sxs-lookup"><span data-stu-id="a3291-221">It already has code that ensures that Promises are supported, even in Internet Explorer 11, and an `Office.onReady` call to assign a handler to the add-in's only button.</span></span>

    > [!NOTE]
    > <span data-ttu-id="a3291-222">名前が示すように、ssoAuthES6.js は JavaScript ES6 構文を使用します。これは、これは、`async`と`await`の使用こそが SSO API の本質的なシンプルさを最もよく示すためです。</span><span class="sxs-lookup"><span data-stu-id="a3291-222">As the name suggests, the ssoAuthES6.js uses JavaScript ES6 syntax because using `async` and `await` best shows the essential simplicity of the SSO API.</span></span> <span data-ttu-id="a3291-223">localhost サーバーが起動するとこのファイルは ES5 構文に変換され、サンプルが Internet Explorer 11 で実行されます。</span><span class="sxs-lookup"><span data-stu-id="a3291-223">When the localhost server is started, this file is transpiled to ES5 syntax so that the sample will run in Internet Explorer 11.</span></span> 

1. <span data-ttu-id="a3291-224">Office.onReady メソッドの下に以下のコードを追加します。</span><span class="sxs-lookup"><span data-stu-id="a3291-224">Add the following code below the Office.onReady method:</span></span>

    ```javascript
    async function getGraphData() {
        try {
            
            // TODO 1: Tell Office to get a bootstrap token from Azure AD.
            
            // TODO 2: Attempt to exchange the bootstrap token for an 
            //         access token to Microsoft Graph.

            // TODO 3: Handle case where Microsoft Graph requires an 
            //         additional form of authentication.

            // TODO 4: Use the access token in a call to Microsoft Graph 
            //         or handle any error from the attempted token exchange.

        }
        catch(exception) {

            // TODO 5: Respond to exceptions thrown by the
            //         OfficeRuntime.auth.getAccessToken call.

        }
    }
    ```

1. <span data-ttu-id="a3291-225">`TODO 1` を次のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="a3291-225">Replace `TODO 1` with the following code.</span></span> <span data-ttu-id="a3291-226">このコードの注意点は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="a3291-226">About this code, note:</span></span>

    - <span data-ttu-id="a3291-227">`OfficeRuntime.auth.getAccessToken`は、Azure AD からブートストラップ トークンを取得するよう Office に指示します。</span><span class="sxs-lookup"><span data-stu-id="a3291-227">`OfficeRuntime.auth.getAccessToken` instructs Office to get a bootstrap token from Azure AD.</span></span> <span data-ttu-id="a3291-228">ブートストラップ トークンは ID トークンに似ていますが、`scp` (スコープ) プロパティ (値`access-as-user`を持つ) を持っています。</span><span class="sxs-lookup"><span data-stu-id="a3291-228">A bootstrap token is similar to an ID token, but it has a `scp` (scope) property with the value `access-as-user`.</span></span> <span data-ttu-id="a3291-229">この種のトークンは、Web アプリケーションによって Microsoft Graph へのアクセス トークンと交換できます。</span><span class="sxs-lookup"><span data-stu-id="a3291-229">This kind of token can be exchanged by a web application for an access token to Microsoft Graph.</span></span>
    - <span data-ttu-id="a3291-230">この `allowSignInPrompt` オプションを true に設定すると、現在 office にサインインしているユーザーがいない場合、office はポップアッププロンプトを開きます。</span><span class="sxs-lookup"><span data-stu-id="a3291-230">Setting the `allowSignInPrompt` option to true means that if no user is currently signed into Office, then Office will open a popup sign-in prompt.</span></span>
    - <span data-ttu-id="a3291-231">この `allowConsentPrompt` オプションを true に設定すると、ユーザーがアドインにユーザーの AAD プロファイルへのアクセスを許可していない場合、Office は同意プロンプトを開きます同意。</span><span class="sxs-lookup"><span data-stu-id="a3291-231">Setting the `allowConsentPrompt` option to true means that if the user has not consented to let the add-in access the user's AAD profile, then Office will open a consent prompt.</span></span> <span data-ttu-id="a3291-232">(このプロンプトでは、ユーザーは、Microsoft Graph の範囲ではなく、ユーザーの AAD プロファイルに同意することのみが許可されます)。</span><span class="sxs-lookup"><span data-stu-id="a3291-232">(The prompt only allows the user to consent to the user's AAD profile, not to Microsoft Graph scopes.)</span></span>
    - <span data-ttu-id="a3291-233">このオプションを true に設定すると、 `forMSGraphAccess` アドインがブートストラップトークンを使用して、ID トークンとして使用するのではなく、Microsoft Graph へのアクセストークンを取得することを Office に通知します。</span><span class="sxs-lookup"><span data-stu-id="a3291-233">Setting the `forMSGraphAccess` option to true signals to Office that the add-in intends to use the bootstrap token to get an access token to Microsoft Graph, instead of just using it as an ID token.</span></span> <span data-ttu-id="a3291-234">テナント管理者が Microsoft Graph へのアドインのアクセスに同意していない場合、`OfficeRuntime.auth.getAccessToken`はエラー **13012** を返します。</span><span class="sxs-lookup"><span data-stu-id="a3291-234">If the tenant administrator has not granted consent to the add-in's access to Microsoft Graph, then `OfficeRuntime.auth.getAccessToken` returns error **13012**.</span></span> <span data-ttu-id="a3291-235">アドインは、Office が Microsoft Graph スコープではなく、ユーザーの Azure AD プロファイルへの同意のみを要求できるために必要となる承認の代替システムにフォールバックすることで応答できます。</span><span class="sxs-lookup"><span data-stu-id="a3291-235">The add-in can respond by falling back to an alternative system of authorization, which is necessary because Office can prompt only for consent to the user's Azure AD profile, not to any Microsoft Graph scopes.</span></span> <span data-ttu-id="a3291-236">フォールバック認証システムでは、ユーザーが再度サインインする必要があり、ユーザーは Microsoft Graph スコープへの同意を求めるメッセージを表示する *ことができ* ます。</span><span class="sxs-lookup"><span data-stu-id="a3291-236">The fallback authorization system requires the user to sign in again and the user *can* be prompted to consent to Microsoft Graph scopes.</span></span> <span data-ttu-id="a3291-237">そのため`forMSGraphAccess`オプションは、同意の欠如により失敗するトークン交換をアドインが行わないことを保証します。</span><span class="sxs-lookup"><span data-stu-id="a3291-237">So, the `forMSGraphAccess` option ensures that the add-in won't make a token exchange that will fail due to lack of consent.</span></span> <span data-ttu-id="a3291-238">(前のステップで管理者の同意が与えられているため、このアドインにおいてはこのシナリオは発生しません。</span><span class="sxs-lookup"><span data-stu-id="a3291-238">(Since you granted administrator consent in an earlier step, this scenario won't happen for this add-in.</span></span> <span data-ttu-id="a3291-239">ベスト プラクティスを示すことを目的として、このオプションはここに含まれています。)</span><span class="sxs-lookup"><span data-stu-id="a3291-239">But the option is included here anyway to illustrate a best practice.)</span></span>

    ```javascript
    let bootstrapToken = await OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: true, allowConsentPrompt: true, forMSGraphAccess: true }); 
    ```

1. <span data-ttu-id="a3291-240">`TODO 2`を以下のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="a3291-240">Replace `TODO 2` with the following code.</span></span> <span data-ttu-id="a3291-241">`getGraphToken`メソッドは後の手順で作成します。</span><span class="sxs-lookup"><span data-stu-id="a3291-241">You'll create the `getGraphToken` method in a later step.</span></span>

    ```javascript
    let exchangeResponse = await getGraphToken(bootstrapToken);
    ```

1. <span data-ttu-id="a3291-242">`TODO 3`を以下のように置き換えます。</span><span class="sxs-lookup"><span data-stu-id="a3291-242">Replace `TODO 3` with the following.</span></span> <span data-ttu-id="a3291-243">このコードについては、以下の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="a3291-243">About this code, note:</span></span> 

    - <span data-ttu-id="a3291-244">Microsoft 365 テナントが多要素認証を要求するように構成されている場合、には、 `exchangeResponse` `claims` 追加の必要な要素に関する情報を含むプロパティが含まれます。</span><span class="sxs-lookup"><span data-stu-id="a3291-244">If the Microsoft 365 tenant has been configured to require multifactor authentication, then the `exchangeResponse` will include a `claims` property with information about the additional required factors.</span></span> <span data-ttu-id="a3291-245">その場合は`OfficeRuntime.auth.getAccessToken`を再度呼び出し、`authChallenge`オプションを Claims プロパティの値に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="a3291-245">In that case, `OfficeRuntime.auth.getAccessToken` should be called again with the `authChallenge` option set to the value of the claims property.</span></span> <span data-ttu-id="a3291-246">これにより、必要なすべての認証形式をユーザーに求めるよう AAD に指示します。</span><span class="sxs-lookup"><span data-stu-id="a3291-246">This tells AAD to prompt the user for all required forms of authentication.</span></span>

    ```javascript
    if (exchangeResponse.claims) {
        let mfaBootstrapToken = await OfficeRuntime.auth.getAccessToken({ authChallenge: exchangeResponse.claims });
        exchangeResponse = await getGraphToken(mfaBootstrapToken);
    }
    ```

1. <span data-ttu-id="a3291-247">`TODO 4`を以下のように置き換えます。</span><span class="sxs-lookup"><span data-stu-id="a3291-247">Replace `TODO 4` with the following.</span></span> <span data-ttu-id="a3291-248">このコードについては、以下の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="a3291-248">About this code, note:</span></span> 

    - <span data-ttu-id="a3291-249">`handleAADErrors`メソッドは後の手順で作成します。</span><span class="sxs-lookup"><span data-stu-id="a3291-249">You'll create the `handleAADErrors` method in a later step.</span></span> <span data-ttu-id="a3291-250">Azure AD エラーは、HTTP コード 200 応答としてクライアントに返されます。</span><span class="sxs-lookup"><span data-stu-id="a3291-250">Azure AD errors are returned to the client as HTTP code 200 Responses.</span></span> <span data-ttu-id="a3291-251">エラーがスローされないため、`catch`ブロック (`getGraphData`メソッドのもの) をトリガーしません。</span><span class="sxs-lookup"><span data-stu-id="a3291-251">They do not throw errors, so they do not trigger the `catch` block of the `getGraphData` method.</span></span>
    - <span data-ttu-id="a3291-252">`makeGraphApiCall`メソッドは後の手順で作成します。</span><span class="sxs-lookup"><span data-stu-id="a3291-252">You'll create the `makeGraphApiCall` method in a later step.</span></span> <span data-ttu-id="a3291-253">これが MS Graph エンドポイントへの AJAX 呼び出しを行います。</span><span class="sxs-lookup"><span data-stu-id="a3291-253">It makes an AJAX call to the MS Graph endpoint.</span></span> <span data-ttu-id="a3291-254">エラーはその呼び出しの`.fail`コールバックでキャッチされます。`catch`ブロック (`getGraphData`メソッドのもの) ではありません。</span><span class="sxs-lookup"><span data-stu-id="a3291-254">Errors are caught in the `.fail` callback of that call, not in the `catch` block of the `getGraphData` method.</span></span>

    ```javascript
    if (exchangeResponse.error) {
        handleAADErrors(exchangeResponse);
    } 
    else {
        makeGraphApiCall(exchangeResponse.access_token);
    }
    ```

1. <span data-ttu-id="a3291-255">`TODO 5`を以下のように置き換えます</span><span class="sxs-lookup"><span data-stu-id="a3291-255">Replace `TODO 5` with the following</span></span>

    - <span data-ttu-id="a3291-256">`getAccessToken`の呼び出しからのエラーは、通常 13xxx の範囲のエラー番号を持つ`code`プロパティを持ちます。</span><span class="sxs-lookup"><span data-stu-id="a3291-256">Errors from the call of `getAccessToken` will have a `code` property with an error number, typically in the 13xxx range.</span></span> <span data-ttu-id="a3291-257">`handleClientSideErrors`メソッドは後の手順で作成します。</span><span class="sxs-lookup"><span data-stu-id="a3291-257">You'll create the `handleClientSideErrors` method in a later step.</span></span>
    - <span data-ttu-id="a3291-258">`showMessage`メソッドは、タスク ウィンドウにテキストを表示します。</span><span class="sxs-lookup"><span data-stu-id="a3291-258">The `showMessage` method displays text on the task pane.</span></span>

    ```javascript
    if (exception.code) { 
        handleClientSideErrors(exception);
    }
    else {
        showMessage("EXCEPTION: " + JSON.stringify(exception));
    }
    ```

1. <span data-ttu-id="a3291-259">`getGraphData`メソッドの下に、以下の関数を追加します。</span><span class="sxs-lookup"><span data-stu-id="a3291-259">Below the `getGraphData` method, add the following function.</span></span> <span data-ttu-id="a3291-260">これ `/auth` は、Microsoft Graph へのアクセストークンについて、ブートストラップトークンを AZURE AD と交換するサーバー側 Express ルートであることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="a3291-260">Note that `/auth` is a server-side Express route that exchanges the bootstrap token with Azure AD for an access token to Microsoft Graph.</span></span>

    ```javascript
    async function getGraphToken(bootstrapToken) {
        let response = await $.ajax({type: "GET", 
            url: "/auth",
            headers: {"Authorization": "Bearer " + bootstrapToken }, 
            cache: false
        });
        return response;
    }
    ```

1. <span data-ttu-id="a3291-261">`getGraphToken`メソッドの下に、以下の関数を追加します。</span><span class="sxs-lookup"><span data-stu-id="a3291-261">Below the `getGraphToken` method, add the following function.</span></span> <span data-ttu-id="a3291-262">`error.code`は数値であり、通常は 13xxx の範囲にあることを注意してください。</span><span class="sxs-lookup"><span data-stu-id="a3291-262">Note that `error.code` is a number, usually in the range 13xxx.</span></span>

    ```javascript
    function handleClientSideErrors(error) {
        switch (error.code) {

            // TODO 6: Handle errors where the add-in should NOT invoke 
            //         the alternative system of authorization.

            // TODO 7: Handle errors where the add-in should invoke 
            //         the alternative system of authorization.

        }
    }
    ```

1. <span data-ttu-id="a3291-263">`TODO 6`を以下のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="a3291-263">Replace `TODO 6` with the following code.</span></span> <span data-ttu-id="a3291-264">これらのエラーの詳細については、「[Office アドインの SSO のトラブルシューティング (Troubleshoot SSO in Office Add-ins)](troubleshoot-sso-in-office-add-ins.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="a3291-264">For more information about these errors, see [Troubleshoot SSO in Office Add-ins](troubleshoot-sso-in-office-add-ins.md).</span></span> 

    ```javascript
    case 13001:
        // No one is signed into Office. If the add-in cannot be effectively used when no one 
        // is logged into Office, then the first call of getAccessToken should pass the 
        // `allowSignInPrompt: true` option. Since this add-in does that, you should not see
        // this error. 
        showMessage("No one is signed into Office. But you can use many of the add-ins functions anyway. If you want to log in, press the Get OneDrive File Names button again.");  
        break;
    case 13002:
        // OfficeRuntime.auth.getAccessToken was called with the allowConsentPrompt 
        // option set to true. But, the user aborted the consent prompt. 
        showMessage("You can use many of the add-ins functions even though you have not granted consent. If you want to grant consent, press the Get OneDrive File Names button again."); 
        break;
    case 13006:
        // Only seen in Office on the web.
        showMessage("Office on the web is experiencing a problem. Please sign out of Office, close the browser, and then start again."); 
        break;
    case 13008:
        // The OfficeRuntime.auth.getAccessToken method has already been called and 
        // that call has not completed yet. Only seen in Office on the web.
        showMessage("Office is still working on the last operation. When it completes, try this operation again."); 
        break;
    case 13010:
        // Only seen in Office on the web.
        showMessage("Follow the instructions to change your browser's zone configuration.");
        break;
    ```

1. <span data-ttu-id="a3291-265">`TODO 7`を以下のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="a3291-265">Replace `TODO 7` with the following code.</span></span> <span data-ttu-id="a3291-266">これらのエラーの詳細については、「[Office アドインの SSO のトラブルシューティング (Troubleshoot SSO in Office Add-ins)](troubleshoot-sso-in-office-add-ins.md)」を参照してください。関数`dialogFallback`は、代替の認証システムを呼び出します。</span><span class="sxs-lookup"><span data-stu-id="a3291-266">For more information about these errors, see [Troubleshoot SSO in Office Add-ins](troubleshoot-sso-in-office-add-ins.md). The function `dialogFallback` invokes the alternative system of authorization.</span></span> <span data-ttu-id="a3291-267">このアドインでは、フォールバック システムはユーザーが既にログインしている場合でもユーザーのサインインを要求するダイアログを開き、msal.js および Implicit Flow を使用して Microsoft Graph へのアクセス トークンを取得します。</span><span class="sxs-lookup"><span data-stu-id="a3291-267">In this add-in, the fallback system opens a dialog which requires the user to sign in, even if the user already is, and uses msal.js and the Implicit Flow to get an access token to Microsoft Graph.</span></span>

    ```javascript
    default:
    // For all other errors, including 13000, 13003, 13005, 13007, 13012, 
    // and 50001, fall back to non-SSO sign-in.
    dialogFallback();
    break;
    ```

1. <span data-ttu-id="a3291-268">`handleClientSideErrors`関数の下に、次の関数を追加します。</span><span class="sxs-lookup"><span data-stu-id="a3291-268">Below the `handleClientSideErrors` function, add the following function.</span></span> 

    ```javascript
    function handleAADErrors(exchangeResponse) {

    // TODO 8: Handle case where the bootstrap token is expired.

    // TODO 9: Handle all other Azure AD errors.
    
    }
    ```

1. <span data-ttu-id="a3291-269">まれに Office がキャッシュしたブートストラップ トークンが Office の検証時に期限切れにならず、交換のために Azure AD に到達するまでの間に期限切れになることがあります。</span><span class="sxs-lookup"><span data-stu-id="a3291-269">On rare occasions the bootstrap token that Office has cached is unexpired when Office validates it, but expires by the time it reaches Azure AD for exchange.</span></span> <span data-ttu-id="a3291-270">Azure AD はエラー **AADSTS500133** で応答します。</span><span class="sxs-lookup"><span data-stu-id="a3291-270">Azure AD will respond with error **AADSTS500133**.</span></span> <span data-ttu-id="a3291-271">この場合、アドインは単に`getGraphData`を再帰的に呼び出す必要があります。</span><span class="sxs-lookup"><span data-stu-id="a3291-271">In this case, the add-in should simply recursively call `getGraphData`.</span></span> <span data-ttu-id="a3291-272">キャッシュされたブートストラップ トークンの有効期限が切れているため、Office は Azure AD から新しいものを取得します。</span><span class="sxs-lookup"><span data-stu-id="a3291-272">Since the cached bootstrap token is now expired, Office will get a new one from Azure AD.</span></span> <span data-ttu-id="a3291-273">そして、`TODO 8`を以下のように置き換えます。</span><span class="sxs-lookup"><span data-stu-id="a3291-273">So replace `TODO 8` with the following.</span></span> 

    ```javascript
    if (exchangeResponse.error_description.indexOf("AADSTS500133") !== -1)
    {
        getGraphData();
    }
    ```

1. <span data-ttu-id="a3291-274">アドインが`getGraphData`の呼び出しの無限ループに入らないようにするため、アドインは`getGraphData`が呼び出された回数を追跡し、1 回以上再帰的に呼び出されないことを確認する必要があります。</span><span class="sxs-lookup"><span data-stu-id="a3291-274">To ensure that the add-in doesn't enter an infinite loop of calls to `getGraphData`, the add-in should keep track of how many times `getGraphData` has been called and be sure that is not called recursively called more than once.</span></span> <span data-ttu-id="a3291-275">そのため、`handleAADErrors`および`getGraphData`関数に対してグローバルなスコープにカウンター変数を作成します。</span><span class="sxs-lookup"><span data-stu-id="a3291-275">So, create a counter variable in a scope that is global to the `handleAADErrors` and `getGraphData` functions.</span></span> <span data-ttu-id="a3291-276">グローバル変数の適切な場所は、`Office.onReady`メソッド呼び出しのすぐ下です。</span><span class="sxs-lookup"><span data-stu-id="a3291-276">A good place for global variables is just below the `Office.onReady` method call.</span></span>

    ```javascript
    let retryGetAccessToken = 0;
    ```

1. <span data-ttu-id="a3291-277">`if`構造 (`handleAADErrors`メソッドのもの) を次のように変更します。</span><span class="sxs-lookup"><span data-stu-id="a3291-277">Change the `if` structure in the `handleAADErrors` method so that it:</span></span>

    - <span data-ttu-id="a3291-278">`getGraphData`を呼び出す直前にカウンターをインクリメントします。</span><span class="sxs-lookup"><span data-stu-id="a3291-278">Increments the counter just before it calls `getGraphData`.</span></span>
    - <span data-ttu-id="a3291-279">`getGraphData`が 2 回目に呼び出されていないことをテストして確認します。</span><span class="sxs-lookup"><span data-stu-id="a3291-279">Tests to ensure that `getGraphData` has not already been called a second time.</span></span> 

    <span data-ttu-id="a3291-280">したがって、`if`構造の最終バージョンは以下のようになります。</span><span class="sxs-lookup"><span data-stu-id="a3291-280">So the final version of the `if` structure should look like the following:</span></span>

    ```javascript
    if ((exchangeResponse.error_description.indexOf("AADSTS500133") !== -1)
        &&
        (retryGetAccessToken <= 0)) 
    {
        retryGetAccessToken++;
        getGraphData();
    }
    ```

1. <span data-ttu-id="a3291-281">`TODO 9`を以下のように置き換えます。</span><span class="sxs-lookup"><span data-stu-id="a3291-281">Replace `TODO 9` with the following.</span></span> 

    ```javascript
    else {
        dialogFallback();
    }
    ```

1. <span data-ttu-id="a3291-282">ファイルを保存して閉じます。</span><span class="sxs-lookup"><span data-stu-id="a3291-282">Save and close the file.</span></span>

### <a name="get-the-data-and-add-it-to-the-office-document"></a><span data-ttu-id="a3291-283">データを取得し、Office ドキュメントへと追加する</span><span class="sxs-lookup"><span data-stu-id="a3291-283">Get the data and add it to the Office document</span></span>

1. <span data-ttu-id="a3291-284">`public\javascripts`フォルダーに、`data.js`という名前の新しいファイルを作成します。</span><span class="sxs-lookup"><span data-stu-id="a3291-284">In the `public\javascripts` folder, create a new file named `data.js`.</span></span>

1. <span data-ttu-id="a3291-285">次の関数をファイルに追加します。</span><span class="sxs-lookup"><span data-stu-id="a3291-285">Add the following function to the file.</span></span> <span data-ttu-id="a3291-286">これは、Microsoft Graph へのアクセス トークンを取得したときに`getGraphData`関数によって呼び出される関数です。  </span><span class="sxs-lookup"><span data-stu-id="a3291-286">This is the function that is called by the `getGraphData` function when it has acquired an access token to Microsoft Graph.</span></span> 

    ```javascript
    function makeGraphApiCall(accessToken) {
        $.ajax(

            // TODO 10: Call an Express route on the add-in's server-side 
            //          code and pass the access token to Microsoft Graph.

        )
        .done(function (response) {

            // TODO 11: Write the data received from Microsoft Graph to 
            //          the Office document.

        })
        .fail(function (errorResult) {
            showMessage("Error from Microsoft Graph: " + JSON.stringify(errorResult));
        });
    }
    ```

1. <span data-ttu-id="a3291-287">`TODO 10`を以下のように置き換えます。</span><span class="sxs-lookup"><span data-stu-id="a3291-287">Replace `TODO 10` with the following.</span></span> <span data-ttu-id="a3291-288">このコードについては、以下の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="a3291-288">About this code, note:</span></span> 

    - <span data-ttu-id="a3291-289">このオブジェクトは、`$.ajax`メソッドのパラメーターです。</span><span class="sxs-lookup"><span data-stu-id="a3291-289">This object is the parameter to the `$.ajax` method.</span></span>
    - <span data-ttu-id="a3291-290">`/getuserdata`は、後の手順で作成するアドインのサーバー上のエクスプレス ルートです。</span><span class="sxs-lookup"><span data-stu-id="a3291-290">The `/getuserdata` is an Express route on the add-in's server that you create in a later step.</span></span> <span data-ttu-id="a3291-291">Microsoft Graph エンドポイントを呼び出し、その呼び出しにアクセス トークンを含めます。</span><span class="sxs-lookup"><span data-stu-id="a3291-291">It will call a Microsoft Graph endpoint and include the access token in its call.</span></span> 

    ```javascript
    {
        type: "GET",
        url: "/getuserdata",
        headers: {"access_token": accessToken },
        cache: false
    }
    ```

1. <span data-ttu-id="a3291-292">`TODO11`を以下のように置き換えます。</span><span class="sxs-lookup"><span data-stu-id="a3291-292">Replace `TODO11` with the following.</span></span> <span data-ttu-id="a3291-293">このコードについては、以下の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="a3291-293">About this code, note:</span></span>

    - <span data-ttu-id="a3291-294">`writeFileNamesToOfficeDocument`は、Graph から Office ドキュメントにデータを挿入します。</span><span class="sxs-lookup"><span data-stu-id="a3291-294">The `writeFileNamesToOfficeDocument` will insert the data from Graph into the Office document.</span></span> <span data-ttu-id="a3291-295">`public\javascripts\document.js`ファイルで定義されています。</span><span class="sxs-lookup"><span data-stu-id="a3291-295">It is defined in the `public\javascripts\document.js` file.</span></span> 
    - <span data-ttu-id="a3291-296">`writeFileNamesToOfficeDocument`がエラーを返した場合、エラー メッセージは "ドキュメントにファイル名を追加できません" で始まります。</span><span class="sxs-lookup"><span data-stu-id="a3291-296">If `writeFileNamesToOfficeDocument` returns an error, it will begin with "Unable to add filenames to document."</span></span>

    ```javascript
    writeFileNamesToOfficeDocument(response)
    .then(function () {
        showMessage("Your data has been added to the document.");
    })
    .catch(function (error) {
        showMessage(error);
    });
    ```

1. <span data-ttu-id="a3291-297">ファイルを保存して閉じます。</span><span class="sxs-lookup"><span data-stu-id="a3291-297">Save and close the file.</span></span>

## <a name="code-the-server-side"></a><span data-ttu-id="a3291-298">サーバー側のコーディング</span><span class="sxs-lookup"><span data-stu-id="a3291-298">Code the server-side</span></span>

### <a name="create-the-auth-router-and-the-token-exchange-logic"></a><span data-ttu-id="a3291-299">認証ルーターおよびトークン交換ロジックを作成する</span><span class="sxs-lookup"><span data-stu-id="a3291-299">Create the auth router and the token exchange logic</span></span>

1. <span data-ttu-id="a3291-300">ファイル`routes\authRoute.js`を開き、`require`ステートメントのすぐ下と`module.exports`ステートメントの上に以下のルート関数を追加します。</span><span class="sxs-lookup"><span data-stu-id="a3291-300">Open the file `routes\authRoute.js` and add the following route function just below the `require` statements and above the `module.exports` statement.</span></span> <span data-ttu-id="a3291-301">`router.get`の URL パラメーターが '/' であることにご注意ください。</span><span class="sxs-lookup"><span data-stu-id="a3291-301">Note that the URL parameter of `router.get` is '/'.</span></span> <span data-ttu-id="a3291-302">このルートは URL '/auth' へのすべての HTTP リクエストを処理するルーターで定義されているため、'/auth' へのすべてのリクエストを効率的に処理します。</span><span class="sxs-lookup"><span data-stu-id="a3291-302">Since this route is being defined in a router that will handle all HTTP Requests for the URL '/auth', this route effectively handles all requests for '/auth'.</span></span> <span data-ttu-id="a3291-303">以前作成したクライアント側の`getGraphToken`関数が、このルートを呼び出します。</span><span class="sxs-lookup"><span data-stu-id="a3291-303">The client-side `getGraphToken` function that you created earlier calls this route.</span></span>  

    ```javascript
    router.get('/', async function(req, res, next) {

        // TODO 12: Test for the presence of the Authorization header.

        // TODO 13: Create the hidden form that will be sent to Azure AD 
        //          to request the access token in exchange for the 
        //          bootstrap token.

        // TODO 14: Send the POST request to Azure AD and relay the 
        //          access token (or an error) to the client.

    });
    ```

1. <span data-ttu-id="a3291-304">`TODO 12`を以下のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="a3291-304">Replace `TODO 12` with the following code.</span></span>

    ```javascript
    const authorization = req.get('Authorization');
    if (authorization == null) {
        let error = new Error('No Authorization header was found.');
        next(error);
    } 
    ```

1. <span data-ttu-id="a3291-305">`TODO 13` を次のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="a3291-305">Replace `TODO 13` with the following code.</span></span> <span data-ttu-id="a3291-306">このコードの注意点は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="a3291-306">About this code, note:</span></span> 

    - <span data-ttu-id="a3291-307">これは長い`else`ブロックの始まりですが、さらにコードを追加するため、終了`}`はまだ終わりではありません。</span><span class="sxs-lookup"><span data-stu-id="a3291-307">This is the beginning of a long `else` block, but the closing `}` is not at the end yet because you will be adding more code to it.</span></span> 
    - <span data-ttu-id="a3291-308">`authorization`文字列は "ベアラー" の後にブートストラップ トークンが続くため、`else`ブロックの最初の行はトークンを`jwt`に割り当てています。</span><span class="sxs-lookup"><span data-stu-id="a3291-308">The `authorization` string is "Bearer " followed by the bootstrap token, so the first line of the `else` block is assigning the token to the `jwt`.</span></span> <span data-ttu-id="a3291-309">("JWT" は "JSON Web Token" の略です)。</span><span class="sxs-lookup"><span data-stu-id="a3291-309">("JWT" stands for "JSON Web Token".)</span></span>
    - <span data-ttu-id="a3291-310">2 つの`process.env.*`値は、アドインを構成したときに割り当てた定数です。</span><span class="sxs-lookup"><span data-stu-id="a3291-310">The two `process.env.*` values are the constants that you assigned when you configured the add-in.</span></span> 
    - <span data-ttu-id="a3291-311">`requested_token_use` フォーム パラメーターは 'on_behalf_of' に設定されています。</span><span class="sxs-lookup"><span data-stu-id="a3291-311">The `requested_token_use` form parameter is set to 'on_behalf_of'.</span></span> <span data-ttu-id="a3291-312">これにより、アドインが On-Behalf-Of フローを使用して Microsoft Graph へのアクセス トークンを要求していることが Azure AD に通知されます。</span><span class="sxs-lookup"><span data-stu-id="a3291-312">This tells Azure AD that the add-in is requesting an access token to Microsoft Graph using the On-Behalf-Of Flow.</span></span> <span data-ttu-id="a3291-313">Azure は、`assertion`フォーム パラメーターに割り当てられているブートストラップ トークンが`scp`プロパティを`access-as-user`に設定された状態で持っていることを検証することで応答します。</span><span class="sxs-lookup"><span data-stu-id="a3291-313">Azure will respond by validating that the bootstrap token, which is assigned to `assertion` form parameter, has a `scp` property that is set to `access-as-user`.</span></span>
    - <span data-ttu-id="a3291-314">`scope`フォーム パラメーターは、アドインが必要とする唯一の Microsoft Graph スコープである 'Files.Read.All' に設定されます。</span><span class="sxs-lookup"><span data-stu-id="a3291-314">The `scope` form parameter is set to 'Files.Read.All' which is the only Microsoft Graph scope that the add-in needs.</span></span>

    ```javascript
     else {
        const [schema, jwt] = authorization.split(' ');
        const formParams = {
        client_id: process.env.CLIENT_ID,
        client_secret: process.env.CLIENT_SECRET,
        grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
        assertion: jwt,
        requested_token_use: 'on_behalf_of',
        scope: ['Files.Read.All'].join(' ')
        };
    ```

1. <span data-ttu-id="a3291-315">`TODO 14`を`else`ブロックを完成させる以下のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="a3291-315">Replace `TODO 14` with the following code, which completes the `else` block.</span></span> <span data-ttu-id="a3291-316">このコードについては、以下の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="a3291-316">About this code, note:</span></span>

    - <span data-ttu-id="a3291-317">const `tenant`は 'common' に設定されます。これは、アドインを Azure AD に登録したときにアドインをマルチテナントとして構成したためです。 特に**サポートされているアカウントの種類**を**任意の組織のディレクトリ内のアカウントと個人用の Microsoft アカウント (例: Skype、Xbox、Outlook.com)** に設定したときです。</span><span class="sxs-lookup"><span data-stu-id="a3291-317">The const `tenant` is set to 'common' because you configured the add-in as multitenant when you registered it with Azure AD; specifically when you set **Supported account types** to **Accounts in any organizational directory and personal Microsoft accounts (e.g. Skype, Xbox, Outlook.com)**.</span></span> <span data-ttu-id="a3291-318">代わりに、アドインが登録されている同じ Microsoft 365 テナントのアカウントのみをサポートすることを選択した場合、このコードでは `tenant` テナントの GUID に設定します。</span><span class="sxs-lookup"><span data-stu-id="a3291-318">If you had instead chosen to support only accounts in the same Microsoft 365 tenancy where the add-in is registered, then in this code `tenant` would be set to the GUID of the tenant.</span></span> 
    - <span data-ttu-id="a3291-319">POST 要求がエラーにならない場合、Azure AD からの応答は JSON に変換され、クライアントに送信されます。</span><span class="sxs-lookup"><span data-stu-id="a3291-319">If the POST request does not error, then the response from Azure AD is converted to JSON and sent to the client.</span></span> <span data-ttu-id="a3291-320">この JSON オブジェクトには、Azure AD が Microsoft Graph へのアクセス トークンを割り当てた`access_token`プロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="a3291-320">This JSON object has an `access_token` property to which Azure AD has assigned the access token to Microsoft Graph.</span></span>

    ```javascript
        const stsDomain = 'https://login.microsoftonline.com';
        const tenant = 'common';
        const tokenURLSegment = 'oauth2/v2.0/token';

        try {
            const tokenResponse = await fetch(`${stsDomain}/${tenant}/${tokenURLSegment}`, {
                method: 'POST',
                body: form(formParams),
                headers: {
                    'Accept': 'application/json',
                    'Content-Type': 'application/x-www-form-urlencoded'
                }
            });
            const json = await tokenResponse.json();

            res.send(json);
        }
        catch(error) {
            res.status(500).send(error);
        }
    }
    ```

1. <span data-ttu-id="a3291-321">ファイルを保存して閉じます。</span><span class="sxs-lookup"><span data-stu-id="a3291-321">Save and close the file.</span></span>

### <a name="create-the-route-that-will-fetch-the-data-from-microsoft-graph"></a><span data-ttu-id="a3291-322">Microsoft Graph からデータを取得するルートを作成する</span><span class="sxs-lookup"><span data-stu-id="a3291-322">Create the route that will fetch the data from Microsoft Graph</span></span>

1. <span data-ttu-id="a3291-323">プロジェクトのルートにある`app.js`ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="a3291-323">Open the file `app.js` in the root of the project.</span></span> <span data-ttu-id="a3291-324">'/dialog.html' のルートのすぐ下に、以下のルートを追加します。</span><span class="sxs-lookup"><span data-stu-id="a3291-324">Just below the route for '/dialog.html', add the following route.</span></span> <span data-ttu-id="a3291-325">このルートは、以前の手順で作成した`makeGraphApiCall`関数によって呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="a3291-325">This route is called by the `makeGraphApiCall` function that you created in an earlier step.</span></span>

    ```javascript
    app.get('/getuserdata', async function(req, res, next) {
        
        // TODO 15: Send a request to the Microsoft Graph REST endpoint.

        // TODO 16: Trim excess information from the returned data and relay it
        //          to the client.
        
    });
    ```

1. <span data-ttu-id="a3291-326">`TODO 15`を以下のように置き換えます。</span><span class="sxs-lookup"><span data-stu-id="a3291-326">Replace `TODO 15` with the following.</span></span> <span data-ttu-id="a3291-327">このコードについては、以下の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="a3291-327">About this code, note:</span></span>

    - <span data-ttu-id="a3291-328">このルートの呼び出し元である`makeGraphApiCall`は、Microsoft Graph へのアクセス トークンを "access_token" という名前のヘッダーとして HTTP 要求に追加しました。</span><span class="sxs-lookup"><span data-stu-id="a3291-328">The caller of this route, `makeGraphApiCall`, added the access token to Microsoft Graph to the HTTP Request as a header named "access_token".</span></span>
    - <span data-ttu-id="a3291-329">`getGraphData`関数は`msgraph-helper.js`ファイルで定義されています。</span><span class="sxs-lookup"><span data-stu-id="a3291-329">The `getGraphData` function is defined in the `msgraph-helper.js` file.</span></span> <span data-ttu-id="a3291-330">(これは、クライアント側の`getGraphData`関数 (`ssoAuthES6.js`ファイルで定義したもの) とは異なります。)</span><span class="sxs-lookup"><span data-stu-id="a3291-330">(This is not the same function as the client-side `getGraphData` function that you defined in the `ssoAuthES6.js` file.)</span></span>
    - <span data-ttu-id="a3291-331">`queryParamsSegment`の最後のパラメーターはハードコーディングされています。</span><span class="sxs-lookup"><span data-stu-id="a3291-331">The last parameter, for `queryParamsSegment`, is hardcoded.</span></span> <span data-ttu-id="a3291-332">本番環境のアドインでこのコードを再利用し、`queryParamsSegment`の一部がユーザーの入力に由来する場合、レスポンス ヘッダー インジェクション攻撃に使用できないようサニタイズされていることをご確認ください。</span><span class="sxs-lookup"><span data-stu-id="a3291-332">If you reuse this code in a production add-in and any part of `queryParamsSegment` comes from user input, be sure that it is sanitized so that it cannot be used in a Response header injection attack.</span></span>
    - <span data-ttu-id="a3291-333">このコードは、必要なプロパティ ("name") および上位 10 のフォルダー名またはファイル名のみを指定することにより、Microsoft Graph から取得する必要があるデータを最小化します。</span><span class="sxs-lookup"><span data-stu-id="a3291-333">The code minimizes the data that must come from Microsoft Graph by specifying only the property we need ("name") and only the top 10 folder or file names.</span></span>

    ```javascript
    const graphToken = req.get('access_token');
    const graphData = await getGraphData(graphToken, "/me/drive/root/children", "?$select=name&$top=10");
    ```

1. <span data-ttu-id="a3291-334">`TODO 16`を以下のように置き換えます。</span><span class="sxs-lookup"><span data-stu-id="a3291-334">Replace `TODO 16` with the following.</span></span> <span data-ttu-id="a3291-335">このコードについては、以下の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="a3291-335">About this code, note:</span></span>

    - <span data-ttu-id="a3291-336">Microsoft Graph が無効なトークンや期限切れトークンなどのエラーを返した場合、返されたオブジェクトには HTTP ステータス (401 など) に設定されたコード プロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="a3291-336">If Microsoft Graph returns an error, such as invalid or expired token, there will be a code property in the returned object set to a HTTP status (e.g., 401).</span></span> <span data-ttu-id="a3291-337">コードはエラーをクライアントに中継します。</span><span class="sxs-lookup"><span data-stu-id="a3291-337">The code relays the error to the client.</span></span> <span data-ttu-id="a3291-338">`.fail`コールバック (`makeGraphApiCall`のもの) でキャッチされます。</span><span class="sxs-lookup"><span data-stu-id="a3291-338">It will be caught in the `.fail` callback of `makeGraphApiCall`.</span></span>
    - <span data-ttu-id="a3291-339">Microsoft Graph データにはアドインが必要としない OData メタデータおよび eTag が含まれているため、コードはクライアントに送信するファイル名のみを含む新しい配列を作成します。</span><span class="sxs-lookup"><span data-stu-id="a3291-339">Microsoft Graph data includes OData metadata and eTags that the add-in does not need, so the code constructs a new array containing only the file names to send to the client.</span></span>

    ```javascript
    if (graphData.code) {
        next(createError(graphData.code, "Microsoft Graph error: " + JSON.stringify(graphData)));
    }
    else {
        const itemNames = [];
        const oneDriveItems = graphData['value'];
        for (let item of oneDriveItems) {
            itemNames.push(item['name']);
        }

        res.send(itemNames)
    }
    ```

1. <span data-ttu-id="a3291-340">ファイルを保存して閉じます。</span><span class="sxs-lookup"><span data-stu-id="a3291-340">Save and close the file.</span></span>

## <a name="run-the-project"></a><span data-ttu-id="a3291-341">プロジェクトを実行する</span><span class="sxs-lookup"><span data-stu-id="a3291-341">Run the project</span></span>

1. <span data-ttu-id="a3291-342">結果を確認できるように、OneDrive 内にファイルがいくつかあることを確認します。</span><span class="sxs-lookup"><span data-stu-id="a3291-342">Ensure that you have some files in your OneDrive so that you can verify the results.</span></span>

1. <span data-ttu-id="a3291-343">`\Begin`フォルダーのルートでコマンド プロンプトを開きます。</span><span class="sxs-lookup"><span data-stu-id="a3291-343">Open a command prompt in the root of the `\Begin` folder.</span></span> 

1. <span data-ttu-id="a3291-344">コマンド`npm start`を実行します。</span><span class="sxs-lookup"><span data-stu-id="a3291-344">Run the command `npm start`.</span></span> 

1. <span data-ttu-id="a3291-345">アドインを Office アプリケーション (Excel、Word、または PowerPoint) にサイドロードして、テストをする必要があります。</span><span class="sxs-lookup"><span data-stu-id="a3291-345">You need to sideload the add-in into an Office application (Excel, Word, or PowerPoint) to test it.</span></span> <span data-ttu-id="a3291-346">手順はプラットフォームによって異なります。</span><span class="sxs-lookup"><span data-stu-id="a3291-346">The instructions depend on your platform.</span></span> <span data-ttu-id="a3291-347">「[テスト用に Office アドインをサイドロードする](../testing/test-debug-office-add-ins.md#sideload-an-office-add-in-for-testing)」に手順へのリンクがあります。</span><span class="sxs-lookup"><span data-stu-id="a3291-347">There are links to instructions at [Sideload an Office Add-in for Testing](../testing/test-debug-office-add-ins.md#sideload-an-office-add-in-for-testing).</span></span>

1. <span data-ttu-id="a3291-348">Office アプリケーションの **[ホーム]** リボンで **[アドインの表示]** ボタン (**SSO Node.js** グループ内) を選択して、作業ウィンドウ アドインを開きます。</span><span class="sxs-lookup"><span data-stu-id="a3291-348">In the Office application, on the **Home** ribbon, select the **Show Add-in** button in the **SSO Node.js** group to open the task pane add-in.</span></span>

1. <span data-ttu-id="a3291-349">**[OneDrive ファイル名の取得]** ボタンをクリックします。</span><span class="sxs-lookup"><span data-stu-id="a3291-349">Click the **Get OneDrive File Names** button.</span></span> <span data-ttu-id="a3291-350">Microsoft 365 の教育機関または職場のアカウント、あるいは Microsoft アカウントを使用して Office にログインしており、SSO が正常に機能している場合は、OneDrive for Business の最初の10個のファイルとフォルダーの名前が文書に挿入されます。</span><span class="sxs-lookup"><span data-stu-id="a3291-350">If you are logged into Office with either a Microsoft 365 Education or work account, or a Microsoft account, and SSO is working as expected, the first 10 file and folder names in your OneDrive for Business are inserted into the document.</span></span> <span data-ttu-id="a3291-351">(最初に 15 秒程度の時間がかかる場合があります。) ログインしていない、または SSO をサポートしていないシナリオにいる場合、もしくは何らかの理由で SSO が機能していない場合には、ログインするように求められます。</span><span class="sxs-lookup"><span data-stu-id="a3291-351">(It may take as much as 15 seconds the first time.) If you are not logged in, or you are in a scenario that does not support SSO, or SSO is not working for any reason, you will be prompted to log in.</span></span> <span data-ttu-id="a3291-352">ログインすると、ファイル名およびフォルダー名が表示されます。</span><span class="sxs-lookup"><span data-stu-id="a3291-352">After you log in, the file and folder names appear.</span></span>

> [!NOTE]
> <span data-ttu-id="a3291-353">以前に別の ID で Office にサインインしており、その時点で開いていた一部の Office アプリケーションがまだ開いている場合、Office が ID を変更したかのように見えても、確実に ID を変更できていない場合があります。</span><span class="sxs-lookup"><span data-stu-id="a3291-353">If you were previously signed into Office with a different ID, and some Office applications that were open at the time are still open, Office may not reliably change your ID even if it appears to have done so.</span></span> <span data-ttu-id="a3291-354">これが発生すると、Microsoft Graph の呼び出しが失敗するか、以前の ID のデータが返される場合があります。</span><span class="sxs-lookup"><span data-stu-id="a3291-354">If this happens, the call to Microsoft Graph may fail or data from the previous ID may be returned.</span></span> <span data-ttu-id="a3291-355">これを防ぐには、必ず*他のすべての Office アプリケーションを閉じて*から、**[OneDrive ファイル名の取得]** を押してください。</span><span class="sxs-lookup"><span data-stu-id="a3291-355">To prevent this, be sure to *close all other Office applications* before you press **Get OneDrive File Names**.</span></span>
