---
title: シングル サインオンを使用する ASP.NET Office アドインを作成する
description: ''
ms.date: 12/04/2019
localization_priority: Normal
ms.openlocfilehash: d9424b1aa0896f9783e2fb7db4160e97bf87cab5
ms.sourcegitcommit: d15bca2c12732f8599be2ec4b2adc7c254552f52
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/12/2020
ms.locfileid: "41950573"
---
# <a name="create-an-aspnet-office-add-in-that-uses-single-sign-on-preview"></a><span data-ttu-id="888f1-102">シングル サインオンを使用する ASP.NET Office アドインを作成する (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="888f1-102">Create an ASP.NET Office Add-in that uses single sign-on (preview)</span></span>

<span data-ttu-id="888f1-103">ユーザーが Office にサインインしたとき、アドインは同じ資格情報を使用し、再度のサインインを要求することなく、複数のアプリケーションへのアクセスを許可することができます。</span><span class="sxs-lookup"><span data-stu-id="888f1-103">When users are signed in to Office, your add-in can use the same credentials to permit users to access multiple applications without requiring them to sign in a second time.</span></span> <span data-ttu-id="888f1-104">概要については、「[Office アドインで SSO を有効化する](sso-in-office-add-ins.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="888f1-104">For an overview, see [Enable SSO in an Office Add-in](sso-in-office-add-ins.md).</span></span>
<span data-ttu-id="888f1-105">この記事では、Node.js と Express を使用して作成したアドインで、シングル サインオン (SSO) を有効化するプロセスについて手順を追って説明します。</span><span class="sxs-lookup"><span data-stu-id="888f1-105">This article walks you through the process of enabling single sign-on (SSO) in an add-in that is built with Node.js and Express.</span></span>

> [!NOTE]
> <span data-ttu-id="888f1-106">ASP.NET ベースのアドインに関する同様の記事については、「[シングル サインオンを使用する Node.js Office アドインを作成する](create-sso-office-add-ins-nodejs.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="888f1-106">For a similar article about an ASP.NET-based add-in, see [Create a Node.js Office Add-in that uses single sign-on](create-sso-office-add-ins-nodejs.md).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="888f1-107">前提条件</span><span class="sxs-lookup"><span data-stu-id="888f1-107">Prerequisites</span></span>

* <span data-ttu-id="888f1-108">Visual Studio 2019 以降。</span><span class="sxs-lookup"><span data-stu-id="888f1-108">Visual Studio 2019 or later.</span></span>

* [<span data-ttu-id="888f1-109">Office Developer Tools</span><span class="sxs-lookup"><span data-stu-id="888f1-109">Office Developer Tools</span></span>](https://www.visualstudio.com/features/office-tools-vs.aspx)

[!include[additional prerequisites](../includes/sso-tutorial-prereqs.md)]

* <span data-ttu-id="888f1-110">Office 365 サブスクリプションの OneDrive for Business に保存されている少なくともいくつかのファイルおよびフォルダー。</span><span class="sxs-lookup"><span data-stu-id="888f1-110">At least a few files and folders stored on OneDrive for Business in your Office 365 subscription.</span></span>

* <span data-ttu-id="888f1-111">Microsoft Azure サブスクリプション。</span><span class="sxs-lookup"><span data-stu-id="888f1-111">A Microsoft Azure subscription.</span></span> <span data-ttu-id="888f1-112">このアドインには、Azure Active Directory (AD) が必要です。</span><span class="sxs-lookup"><span data-stu-id="888f1-112">This add-in requires Azure Active Directory (AD).</span></span> <span data-ttu-id="888f1-113">Azure AD は、アプリケーションが認証および承認に使用する ID サービスを提供します。</span><span class="sxs-lookup"><span data-stu-id="888f1-113">Azure AD provides identity services that applications use for authentication and authorization.</span></span> <span data-ttu-id="888f1-114">[Microsoft Azure](https://account.windowsazure.com/SignUp) で試用版サブスクリプションを取得できます。</span><span class="sxs-lookup"><span data-stu-id="888f1-114">A trial subscription can be acquired at [Microsoft Azure](https://account.windowsazure.com/SignUp).</span></span>

## <a name="set-up-the-starter-project"></a><span data-ttu-id="888f1-115">スタート プロジェクトをセットアップする</span><span class="sxs-lookup"><span data-stu-id="888f1-115">Set up the starter project</span></span>

<span data-ttu-id="888f1-116">「[Office Add-in ASPNET SSO](https://github.com/officedev/office-add-in-aspnet-sso)」にあるリポジトリを複製するかダウンロードします。</span><span class="sxs-lookup"><span data-stu-id="888f1-116">Clone or download the repo at [Office Add-in ASPNET SSO](https://github.com/officedev/office-add-in-aspnet-sso).</span></span>

> [!NOTE]
> <span data-ttu-id="888f1-117">サンプルには 2 つのバージョンがあります。</span><span class="sxs-lookup"><span data-stu-id="888f1-117">There are two versions of the sample:</span></span>
>
> * <span data-ttu-id="888f1-p103">**[Before]** フォルダーはスタート プロジェクトです。SSO や承認に直接関連しない UI などの側面は、既に完了しています。この記事で後述する各セクションでは、これを完成させるための手順を順に説明します。</span><span class="sxs-lookup"><span data-stu-id="888f1-p103">The **Before** folder is a starter project. The UI and other aspects of the add-in that are not directly connected to SSO or authorization are already done. Later sections of this article walk you through the process of completing it.</span></span>
> * <span data-ttu-id="888f1-121">このサンプルの **[Complete]** バージョンは、この記事の手順を完了したときに得られるアドインと同様のものですが、完成済みのプロジェクトには、この記事のテキストと重複するコード コメントが含まれています。</span><span class="sxs-lookup"><span data-stu-id="888f1-121">The **Complete** version of the sample is just like the add-in that you would have if you completed the procedures of this article, except that the completed project has code comments that would be redundant with the text of this article.</span></span> <span data-ttu-id="888f1-122">完成済みのバージョンを使用する場合は、この記事の手順をそのまま実行しますが、[Before] を [Complete] に置き換えて、「**クライアント側のコードを作成する**」と「**サーバー側のコードを作成する**」のセクションを省略してください。</span><span class="sxs-lookup"><span data-stu-id="888f1-122">To use the completed version, just follow the instructions in this article, but replace "Before" with "Complete" and skip the sections **Code the client side** and **Code the server side**.</span></span>


## <a name="register-the-add-in-with-azure-ad-v20-endpoint"></a><span data-ttu-id="888f1-123">Azure AD v2.0 エンドポイントにアドインを登録する</span><span class="sxs-lookup"><span data-stu-id="888f1-123">Register the add-in with Azure AD v2.0 endpoint</span></span>

1. <span data-ttu-id="888f1-124">[Azure ポータル - アプリの登録](https://go.microsoft.com/fwlink/?linkid=2083908)ページに移動してアプリを登録します。</span><span class="sxs-lookup"><span data-stu-id="888f1-124">Navigate to the [Azure portal - App registrations](https://go.microsoft.com/fwlink/?linkid=2083908) page to register your app.</span></span>

1. <span data-ttu-id="888f1-125">***管理者***の資格情報を使用して Office 365 テナントにサインインします。</span><span class="sxs-lookup"><span data-stu-id="888f1-125">Sign in with the ***admin*** credentials to your Office 365 tenancy.</span></span> <span data-ttu-id="888f1-126">たとえば、MyName@contoso.onmicrosoft.com です。</span><span class="sxs-lookup"><span data-stu-id="888f1-126">For example, MyName@contoso.onmicrosoft.com.</span></span>

1. <span data-ttu-id="888f1-127">**[新規登録]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="888f1-127">Select **New registration**.</span></span> <span data-ttu-id="888f1-128">**[アプリケーションを登録]** ページで、次のように値を設定します。</span><span class="sxs-lookup"><span data-stu-id="888f1-128">On the **Register an application** page, set the values as follows.</span></span>

    * <span data-ttu-id="888f1-129">`Office-Add-in-ASPNET-SSO` に **[名前]** を設定します。</span><span class="sxs-lookup"><span data-stu-id="888f1-129">Set **Name** to `Office-Add-in-ASPNET-SSO`.</span></span>
    * <span data-ttu-id="888f1-130">[**サポートされているアカウントの種類**] を [**任意の組織のディレクトリ内のアカウント (任意の Azure AD ディレクトリ - マルチテナント) と個人用の Microsoft アカウント (例: Skype、 Xbox)**] に設定します。</span><span class="sxs-lookup"><span data-stu-id="888f1-130">Set **Supported account types** to **Accounts in any organizational directory (Any Azure AD directory - Multitenant) and personal Microsoft accounts (e.g. Skype, Xbox)**.</span></span> <span data-ttu-id="888f1-131">(登録しているテナントのユーザーだけがアドインを使用できるようにする場合は、代わりに [**この組織ディレクトリのアカウントのみ...**] を選択します。ただし、追加セットアップ手順を実行する必要があります。</span><span class="sxs-lookup"><span data-stu-id="888f1-131">(If you want the add-in to be usable only by users in the tenancy where you are registering it, you can choose **Accounts in this organizational directory only ...** instead, but you will need to go through some additional setup steps.</span></span> <span data-ttu-id="888f1-132">詳細については、「**シングルテナントのセットアップ**」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="888f1-132">See **Setup for single-tenant** below.)</span></span>
    * <span data-ttu-id="888f1-133">[**リダイレクト URI**] セクションで、ドロップダウンで [**Web**] が選択されていることを確認し、URI を [` https://localhost:44355/AzureADAuth/Authorize`] に設定します。</span><span class="sxs-lookup"><span data-stu-id="888f1-133">In the **Redirect URI** section, ensure that **Web** is selected in the drop down and then set the URI to` https://localhost:44355/AzureADAuth/Authorize`.</span></span>
    * <span data-ttu-id="888f1-134">**[登録]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="888f1-134">Choose **Register**.</span></span>

1. <span data-ttu-id="888f1-135">**Office-Add-in-NodeJS-SSO** ページで、**アプリケーション (クライアント) ID** と**ディレクトリ (テナント) ID** の値をコピーして保存します。</span><span class="sxs-lookup"><span data-stu-id="888f1-135">On the **Office-Add-in-NodeJS-SSO** page, copy and save the values for the **Application (client) ID** and the **Directory (tenant) ID**.</span></span> <span data-ttu-id="888f1-136">以降の手順では、それらの両方を使用します。</span><span class="sxs-lookup"><span data-stu-id="888f1-136">You'll use both of them in later procedures.</span></span>

    > [!NOTE]
    > <span data-ttu-id="888f1-137">この ID は、Office ホスト アプリケーション (たとえば、PowerPoint、Word、Excel) などの別のアプリケーションが、このアプリケーションへの承認されたアクセスを求めるときの「対象ユーザー」値になります。</span><span class="sxs-lookup"><span data-stu-id="888f1-137">This ID is the "audience" value when other applications, such as the Office host application (e.g., PowerPoint, Word, Excel), seek authorized access to the application.</span></span> <span data-ttu-id="888f1-138">また、そのアプリケーションが Microsoft Graph への承認されたアクセスを求めるときには、このアプリケーションの「クライアント ID」になります。</span><span class="sxs-lookup"><span data-stu-id="888f1-138">It is also the "client ID" of the application when it, in turn, seeks authorized access to Microsoft Graph.</span></span>

1. <span data-ttu-id="888f1-139">[**管理**] で [**証明書とシークレット**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="888f1-139">Under **Manage**, select **Certificates & secrets**.</span></span> <span data-ttu-id="888f1-140">[**新しいクライアント シークレット**] ボタンを選択します。</span><span class="sxs-lookup"><span data-stu-id="888f1-140">Select the **New client secret** button.</span></span> <span data-ttu-id="888f1-141">[**説明**] に値を入力してから、[**有効期限**] の適切なオプションを選択し、[**追加**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="888f1-141">Enter a value for **Description**, then select an appropriate option for **Expires** and choose **Add**.</span></span> <span data-ttu-id="888f1-142">後の手順で必要になるため、先に進む前に、*クライアント シークレットの値をすぐにコピーし、アプリケーション ID とともに保存*します。</span><span class="sxs-lookup"><span data-stu-id="888f1-142">*Copy the client secret value immediately and save it with the application ID* before proceeding as you'll need it in a later procedure.</span></span>

1. <span data-ttu-id="888f1-143">[**管理**] で [**API の公開**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="888f1-143">Under **Manage**, select **Expose an API**.</span></span> <span data-ttu-id="888f1-144">**[設定]** リンクを選択して、"api://$App ID GUID$" の形式でアプリケーション ID URI を生成します。$App ID GUID$ は**アプリケーション (クライアント) ID** です。</span><span class="sxs-lookup"><span data-stu-id="888f1-144">Select the **Set** link to generate the Application ID URI in the form "api://$App ID GUID$", where $App ID GUID$ is the **Application (client) ID**.</span></span> <span data-ttu-id="888f1-145">`//` の後と GUID の前に、`localhost:44355/` を挿入します (末尾に追加されたスラッシュ「/」に注意します)。</span><span class="sxs-lookup"><span data-stu-id="888f1-145">Insert `localhost:44355/` (note the forward slash "/" appended to the end) after the `//` and before the GUID.</span></span> <span data-ttu-id="888f1-146">ID 全体の形式は `api://localhost:44355/$App ID GUID$` でなければなりません (例: `api://localhost:44355/c6c1f32b-5e55-4997-881a-753cc1d563b7`)。</span><span class="sxs-lookup"><span data-stu-id="888f1-146">The entire ID should have the form `api://localhost:44355/$App ID GUID$`; for example `api://localhost:44355/c6c1f32b-5e55-4997-881a-753cc1d563b7`.</span></span>

1. <span data-ttu-id="888f1-147">ダイアログで [**保存**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="888f1-147">Select **Save** on the dialog.</span></span>

1. <span data-ttu-id="888f1-148">**[Scope の追加]** ボタンをクリックします。</span><span class="sxs-lookup"><span data-stu-id="888f1-148">Select the **Add a scope** button.</span></span> <span data-ttu-id="888f1-149">開いたパネルで、`access_as_user`を **[スコープ名]** として入力します。</span><span class="sxs-lookup"><span data-stu-id="888f1-149">In the panel that opens, enter `access_as_user` as the **Scope** name.</span></span>

1. <span data-ttu-id="888f1-150">**[同意できるのはだれですか?]** を **[管理者とユーザー]** に設定します。</span><span class="sxs-lookup"><span data-stu-id="888f1-150">Set **Who can consent?** to **Admins and users**.</span></span>

1. <span data-ttu-id="888f1-151">管理者およびユーザーの同意のプロンプトを構成するためのフィールドに、現在のユーザーと同じ権限で Office ホスト アプリケーションがアドインの Web API を使用できるようにする `access_as_user` 範囲に適した値を入力します。</span><span class="sxs-lookup"><span data-stu-id="888f1-151">Fill in the fields for configuring the admin and user consent prompts with values that are appropriate for the `access_as_user` scope which enables the Office host application to use your add-in's web APIs with the same rights as the current user.</span></span> <span data-ttu-id="888f1-152">提案:</span><span class="sxs-lookup"><span data-stu-id="888f1-152">Suggestions:</span></span>

    - <span data-ttu-id="888f1-153">**管理者の同意のタイトル**: Office はユーザーとして機能できます。</span><span class="sxs-lookup"><span data-stu-id="888f1-153">**Admin consent title**: Office can act as the user.</span></span>
    - <span data-ttu-id="888f1-154">**管理者の同意の説明**: 現在のユーザーと同じ権限で Office がアドインの Web API を呼び出すことを可能にします。</span><span class="sxs-lookup"><span data-stu-id="888f1-154">**Admin consent description**: Enable Office to call the add-in's web APIs with the same rights as the current user.</span></span>
    - <span data-ttu-id="888f1-155">**ユーザーの同意のタイトル**: Office は自分として機能できます。</span><span class="sxs-lookup"><span data-stu-id="888f1-155">**User consent title**: Office can act as you.</span></span>
    - <span data-ttu-id="888f1-156">**管理者の同意の説明**: 自分と同じ権限で Office がアドインの Web API を呼び出すことを可能にします。</span><span class="sxs-lookup"><span data-stu-id="888f1-156">**Admin consent description**: Enable Office to call the add-in's web APIs with the same rights that you have.</span></span>

1. <span data-ttu-id="888f1-157">**[状態]** が **[有効]** に設定されていることを確認してください。</span><span class="sxs-lookup"><span data-stu-id="888f1-157">Ensure that **State** is set to **Enabled**.</span></span>

1. <span data-ttu-id="888f1-158">**[スコープの追加]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="888f1-158">Select **Add scope** .</span></span>

    > [!NOTE]
    > <span data-ttu-id="888f1-159">テキストフィールドのすぐ下に表示される **[スコープ名]** のドメイン部分は、以前に設定したアプリケーション ID URI に自動的に一致し、末尾に`/access_as_user`が追加されます。たとえば、`api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`です。</span><span class="sxs-lookup"><span data-stu-id="888f1-159">The domain part of the **Scope** name displayed just below the text field should automatically match the Application ID URI that you set earlier, with `/access_as_user` appended to the end; for example, `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`.</span></span>

1. <span data-ttu-id="888f1-160">**[承認済みのクライアント アプリケーション]** セクションで、アドインの Web アプリケーションに対して承認するアプリケーションを特定します。</span><span class="sxs-lookup"><span data-stu-id="888f1-160">In the **Authorized client applications** section, you identify the applications that you want to authorize to your add-in's web application.</span></span> <span data-ttu-id="888f1-161">次のそれぞれの ID を事前承認する必要があります。</span><span class="sxs-lookup"><span data-stu-id="888f1-161">Each of the following IDs needs to be pre-authorized.</span></span>

    - <span data-ttu-id="888f1-162">`d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)</span><span class="sxs-lookup"><span data-stu-id="888f1-162">`d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)</span></span>
    - <span data-ttu-id="888f1-163">`ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (Microsoft Office)</span><span class="sxs-lookup"><span data-stu-id="888f1-163">`ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (Microsoft Office)</span></span>
    - <span data-ttu-id="888f1-164">`57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office on the web)</span><span class="sxs-lookup"><span data-stu-id="888f1-164">`57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office on the web)</span></span>
    - <span data-ttu-id="888f1-165">`bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Outlook on the web)</span><span class="sxs-lookup"><span data-stu-id="888f1-165">`bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Outlook on the web)</span></span>

    <span data-ttu-id="888f1-166">ID ごとに、次の手順を実行します。</span><span class="sxs-lookup"><span data-stu-id="888f1-166">For each ID, take these steps:</span></span>

    <span data-ttu-id="888f1-167">a. </span><span class="sxs-lookup"><span data-stu-id="888f1-167">a.</span></span> <span data-ttu-id="888f1-168">**[クライアント アプリケーションの追加]** ボタンを選択し、表示されたパネルで [クライアント ID] をそれぞれの GUID に設定して、`api://localhost:44355/$App ID GUID$/access_as_user`のチェック ボックスをオンにします。</span><span class="sxs-lookup"><span data-stu-id="888f1-168">Select **Add a client application** button and then, in the panel that opens, set the Client ID to the respective GUID and check the box for `api://localhost:44355/$App ID GUID$/access_as_user`.</span></span>

    <span data-ttu-id="888f1-169">b. </span><span class="sxs-lookup"><span data-stu-id="888f1-169">b.</span></span> <span data-ttu-id="888f1-170">**[アプリケーションの追加]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="888f1-170">Select **Add application**.</span></span>

1. <span data-ttu-id="888f1-171">[**管理**] で [**API のアクセス許可**]、[**アクセス許可の追加**] の順に選択します。</span><span class="sxs-lookup"><span data-stu-id="888f1-171">Under **Manage**, select **API permissions** and then select **Add a permission**.</span></span> <span data-ttu-id="888f1-172">開いたパネルで、**[Microsoft Graph]** を選択してから **[委任されたアクセス許可]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="888f1-172">On the panel that opens, choose **Microsoft Graph** and then choose **Delegated permissions**.</span></span>

1. <span data-ttu-id="888f1-173">アドインに必要な権限を検索するには、**[アクセス許可を選択]** の検索ボックスを使用します。</span><span class="sxs-lookup"><span data-stu-id="888f1-173">Use the **Select permissions** search box to search for the permissions your add-in needs.</span></span> <span data-ttu-id="888f1-174">以下を選択します。</span><span class="sxs-lookup"><span data-stu-id="888f1-174">Select the following.</span></span> <span data-ttu-id="888f1-175">アドイン自体に実際に必要なものは最初のもののみですが、Office ホストがアドインの Web アプリケーションへのトークンを取得するには、`profile` のアクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="888f1-175">Only the first is really required by your add-in itself; but the `profile` permission is required for the Office host to get a token to your add-in web application.</span></span> <span data-ttu-id="888f1-176">(実際には、Files.Read.All とプロファイルのみがアドインに必要です。</span><span class="sxs-lookup"><span data-stu-id="888f1-176">(Only Files.Read.All and profile are actually needed by the add-in.</span></span> <span data-ttu-id="888f1-177">MSAL.NET ライブラリに必要なので、他の 2 つを要求する必要があります。)</span><span class="sxs-lookup"><span data-stu-id="888f1-177">You must request the other two because the MSAL.NET library requires them.)</span></span>

    * <span data-ttu-id="888f1-178">Files.Read.All</span><span class="sxs-lookup"><span data-stu-id="888f1-178">Files.Read.All</span></span>
    * <span data-ttu-id="888f1-179">offline_access</span><span class="sxs-lookup"><span data-stu-id="888f1-179">offline_access</span></span>
    * <span data-ttu-id="888f1-180">openid</span><span class="sxs-lookup"><span data-stu-id="888f1-180">openid</span></span>
    * <span data-ttu-id="888f1-181">profile</span><span class="sxs-lookup"><span data-stu-id="888f1-181">profile</span></span>

    > [!NOTE]
    > <span data-ttu-id="888f1-182">`User.Read` アクセス許可は既定でリストされています。</span><span class="sxs-lookup"><span data-stu-id="888f1-182">The `User.Read` permission may already be listed by default.</span></span> <span data-ttu-id="888f1-183">必要でないアクセス許可は依頼しない方がよいため、アドインが実際に必要でない場合は、このアクセス許可のボックスのチェックをオフにしておくことをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="888f1-183">It is a good practice not to ask for permissions that are not needed, so we recommend that you uncheck the box for this permission if your add-in does not actually need it.</span></span>

1. <span data-ttu-id="888f1-184">表示される各アクセス許可のチェック ボックスをオンにします。</span><span class="sxs-lookup"><span data-stu-id="888f1-184">Select the check box for each permission as it appears.</span></span> <span data-ttu-id="888f1-185">アドインに必要なアクセス許可を選択したら、パネルの下部にある **[アクセス許可を追加する]** ボタンをクリックします。</span><span class="sxs-lookup"><span data-stu-id="888f1-185">After selecting the permissions that your add-in needs, select the **Add permissions** button at the bottom of the panel.</span></span>

1. <span data-ttu-id="888f1-186">同じページで、[**[テナント名] に管理者の同意を与える**] ボタンを選択し、表示される確認に対して [**同意する**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="888f1-186">On the same page, choose the **Grant admin consent for [tenant name]** button, and then select **Accept** for the confirmation that appears.</span></span>

    > [!NOTE]
    > <span data-ttu-id="888f1-187">[**[テナント名] に管理者の同意を与える**] を選択すると、同意プロンプトを作成できるように、数分後に再試行を求めるバナー メッセージが表示される場合があります。</span><span class="sxs-lookup"><span data-stu-id="888f1-187">After choosing **Grant admin consent for [tenant name]**, you may see a banner message asking you to try again in a few minutes so that the consent prompt can be constructed.</span></span> <span data-ttu-id="888f1-188">その場合は、次のセクションで作業を開始できますが、***必ずポータルに戻り、このボタンを押してください***。</span><span class="sxs-lookup"><span data-stu-id="888f1-188">If so, you can start work on the next section, ***but don't forget to come back to the portal and press this button***!</span></span>

## <a name="configure-the-solution"></a><span data-ttu-id="888f1-189">ソリューションを構成する</span><span class="sxs-lookup"><span data-stu-id="888f1-189">Configure the solution</span></span>

1. <span data-ttu-id="888f1-190">[**Before**] フォルダーのルートで、**Visual Studio** でソリューション (.sln) ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="888f1-190">In the root of the **Before** folder, open the solution (.sln) file in **Visual Studio**.</span></span> <span data-ttu-id="888f1-191">[**ソリューション エクスプローラー**] の一番上のノード (プロジェクト ノードではなく、ソリューション ノード) を右クリックして、[**スタートアップ プロジェクトの設定**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="888f1-191">Right-click the top node in **Solution Explorer** (the Solution node, not either of the project nodes), and then select **Set startup projects**.</span></span>

1. <span data-ttu-id="888f1-192">[**共通プロパティ**] で、[**スタートアップ プロジェクト**]、[**マルチ スタートアップ プロジェクト**] の順に選択します。</span><span class="sxs-lookup"><span data-stu-id="888f1-192">Under **Common Properties**, select **Startup Project**, and then **Multiple startup projects**.</span></span> <span data-ttu-id="888f1-193">両方のプロジェクトの [**アクション**] が [**開始**] に設定され、「... WebAPI」で終わるプロジェクトが最初にリストされていることを確認します。</span><span class="sxs-lookup"><span data-stu-id="888f1-193">Ensure that the **Action** for both projects is set to **Start**, and that the project that ends in "...WebAPI" is listed first.</span></span> <span data-ttu-id="888f1-194">ダイアログを閉じます。</span><span class="sxs-lookup"><span data-stu-id="888f1-194">Close the dialog.</span></span>

1. <span data-ttu-id="888f1-195">[**ソリューション エクスプローラー**] に戻り、[**Office-Add-in-Microsoft-Graph-ASPNETWebAPI**] プロジェクトを選択します (右クリックしないでください)。</span><span class="sxs-lookup"><span data-stu-id="888f1-195">Back in **Solution Explorer**, select (don't right-click) the **Office-Add-in-Microsoft-Graph-ASPNETWebAPI** project.</span></span> <span data-ttu-id="888f1-196">[**プロパティ**] ウィンドウを開きます。</span><span class="sxs-lookup"><span data-stu-id="888f1-196">The **Properties** pane opens.</span></span> <span data-ttu-id="888f1-197">[**SSL 有効**] が [**True**] であることを確認します。</span><span class="sxs-lookup"><span data-stu-id="888f1-197">Ensure that **SSL Enabled** is **True**.</span></span> <span data-ttu-id="888f1-198">[**SSL URL**] が `http://localhost:44355/` であることを確認します。</span><span class="sxs-lookup"><span data-stu-id="888f1-198">Verify that the **SSL URL** is `http://localhost:44355/`.</span></span>

1. <span data-ttu-id="888f1-199">「Web.config」 で、以前にコピーした値を使用します。</span><span class="sxs-lookup"><span data-stu-id="888f1-199">In "Web.config", use the values that you copied in earlier.</span></span> <span data-ttu-id="888f1-200">[**ida:ClientID**] と [**ida:Audience**] の両方を [**アプリケーション (クライアント) ID**] に設定し、[**ida:Password**] をクライアント シークレットに設定します。</span><span class="sxs-lookup"><span data-stu-id="888f1-200">Set both the **ida:ClientID** and the **ida:Audience** to your **Application (client) ID**, and set **ida:Password** to your client secret.</span></span>

    > [!NOTE]
    > <span data-ttu-id="888f1-201">[**アプリケーション (クライアント) ID**] は、Office ホスト アプリケーション (たとえば、PowerPoint、Word、Excel) などの別のアプリケーションが、このアプリケーションへの承認されたアクセスを求めるときの「対象ユーザー」値になります。</span><span class="sxs-lookup"><span data-stu-id="888f1-201">The **Application (client) ID** is the "audience" value when other applications, such as the Office host application (e.g., PowerPoint, Word, Excel), seek authorized access to the application.</span></span> <span data-ttu-id="888f1-202">また、そのアプリケーションが Microsoft Graph への承認されたアクセスを求めるときには、このアプリケーションの「クライアント ID」になります。</span><span class="sxs-lookup"><span data-stu-id="888f1-202">It is also the "client ID" of the application when it, in turn, seeks authorized access to Microsoft Graph.</span></span>

1. <span data-ttu-id="888f1-203">アドインを登録したときに、**サポートされているアカウントの種類**で「この組織のディレクトリ内のアカウントのみ」を選択しなかった場合は、web.config を保存して閉じます。 それ以外の場合は、保存して、開いたままにします。</span><span class="sxs-lookup"><span data-stu-id="888f1-203">If you didn't choose "Accounts in this organizational directory only" for **SUPPORTED ACCOUNT TYPES** when you registered the add-in, save and close the web.config. Otherwise, save but leave it open.</span></span>

1. <span data-ttu-id="888f1-204">[**ソリューション エクスプローラー**] で [**Office-Add-in-Microsoft-Graph-ASPNET**] プロジェクトを選択し、アドイン マニフェスト ファイル「Office-Add-in-ASPNET-SSO.xml」を開いて、ファイルの下部までスクロールします。</span><span class="sxs-lookup"><span data-stu-id="888f1-204">Still in **Solution Explorer**, choose the **Office-Add-in-Microsoft-Graph-ASPNET** project and open the add-in manifest file “Office-Add-in-ASPNET-SSO.xml” and then scroll to the bottom of the file.</span></span> <span data-ttu-id="888f1-205">`</VersionOverrides>` 終了タグの直前に、以下のマークアップがあります。</span><span class="sxs-lookup"><span data-stu-id="888f1-205">Just above the end `</VersionOverrides>` tag, you'll find the following markup:</span></span>

    ```xml
    <WebApplicationInfo>
      <Id>$application_GUID here$</Id>
      <Resource>api://localhost:44355/$application_GUID here$</Resource>
      <Scopes>
          <Scope>Files.Read.All</Scope>
          <Scope>offline_access</Scope>
          <Scope>openid</Scope>
          <Scope>profile</Scope>
      </Scopes>
    </WebApplicationInfo>
    ```

1. <span data-ttu-id="888f1-206">このマークアップ内の*両方の場所の*プレースホルダー「$application_GUID here$」を、アドインの登録時にコピーしたアプリケーション ID に置き換えます。</span><span class="sxs-lookup"><span data-stu-id="888f1-206">Replace the placeholder “$application_GUID here$” *in both places* in the markup with the Application ID that you copied when you registered your add-in.</span></span> <span data-ttu-id="888f1-207">「$」は ID の一部ではないので、これらを含めないでください。</span><span class="sxs-lookup"><span data-stu-id="888f1-207">The "$" signs are not part of the ID, so do not include them.</span></span> <span data-ttu-id="888f1-208">これは、web.config の ClientID と Audience に使用したものと同じ ID です。</span><span class="sxs-lookup"><span data-stu-id="888f1-208">This is the same ID you used in for the ClientID and Audience in the web.config.</span></span>

  > [!NOTE]
  > <span data-ttu-id="888f1-209">**リソース**値は、アドインを登録したときに設定した**アプリケーション ID URI** です。</span><span class="sxs-lookup"><span data-stu-id="888f1-209">The **Resource** value is the **Application ID URI** you set when you registered the add-in.</span></span> <span data-ttu-id="888f1-210">**[範囲]** セクションは、アドインが AppSource を通じて販売される場合に同意ダイアログ ボックスを生成するためにのみ使用されます。</span><span class="sxs-lookup"><span data-stu-id="888f1-210">The **Scopes** section is used only to generate a consent dialog box if the add-in is sold through AppSource.</span></span>

1. <span data-ttu-id="888f1-211">ファイルを保存して閉じます。</span><span class="sxs-lookup"><span data-stu-id="888f1-211">Save and close the file.</span></span>

### <a name="setup-for-single-tenant"></a><span data-ttu-id="888f1-212">シングルテナントのセットアップ</span><span class="sxs-lookup"><span data-stu-id="888f1-212">Setup for single-tenant</span></span>

<span data-ttu-id="888f1-213">アドインを登録したときに、**サポートされているアカウントの種類**で「この組織のディレクトリ内のアカウントのみ」を選択した場合は、これらの追加のセットアップ手順を実行する必要があります。</span><span class="sxs-lookup"><span data-stu-id="888f1-213">If you chose "Accounts in this organizational directory only" for **SUPPORTED ACCOUNT TYPES** when you registered the add-in, you need to take these additional setup steps:</span></span>

1. <span data-ttu-id="888f1-214">Azure ポータルに戻り、アドインの登録の [**概要**] ブレードを開きます。</span><span class="sxs-lookup"><span data-stu-id="888f1-214">Go back to the Azure Portal and open the **Overview** blade of the add-in's registration.</span></span> <span data-ttu-id="888f1-215">[**Directory (テナント) ID**] をコピーします。</span><span class="sxs-lookup"><span data-stu-id="888f1-215">Copy the **Directory (tenant) ID**.</span></span>

1. <span data-ttu-id="888f1-216">web.config で、[**ida：Authority**] の値の「Common」を前の手順でコピーした GUID に置き換えます。</span><span class="sxs-lookup"><span data-stu-id="888f1-216">In the web.config, replace the "common" in the value of **ida:Authority** with the GUID you copied in the preceding step.</span></span> <span data-ttu-id="888f1-217">終了すると、値は `<add key="ida:Authority" value="https://login.microsoftonline.com/12345678-91ab-cdef-0123-456789abcdef/oauth2/v2.0" />` のようになります。</span><span class="sxs-lookup"><span data-stu-id="888f1-217">When you are finished the value should look similar to this: `<add key="ida:Authority" value="https://login.microsoftonline.com/12345678-91ab-cdef-0123-456789abcdef/oauth2/v2.0" />`.</span></span>

1. <span data-ttu-id="888f1-218">web.config を保存して閉じます。</span><span class="sxs-lookup"><span data-stu-id="888f1-218">Save and close the web.config.</span></span>

## <a name="code-the-client-side"></a><span data-ttu-id="888f1-219">クライアント側のコードの作成</span><span class="sxs-lookup"><span data-stu-id="888f1-219">Code the client side</span></span>

1. <span data-ttu-id="888f1-220">[**スクリプト**] フォルダー内の HomeES6.js ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="888f1-220">Open the HomeES6.js file in the **Scripts** folder.</span></span> <span data-ttu-id="888f1-221">これには、一部のコードが既に含まれています。</span><span class="sxs-lookup"><span data-stu-id="888f1-221">It already has some code in it:</span></span>

    * <span data-ttu-id="888f1-222">Office が UI に Internet Explorer を使用しているときにアドインを実行できるように、Office.Promise オブジェクトをグローバル ウィンドウ オブジェクトに割り当てるポリフィル。</span><span class="sxs-lookup"><span data-stu-id="888f1-222">A polyfill that assigns the Office.Promise object to the global window object so that the add-in can run when Office is using Internet Explorer for the UI.</span></span> <span data-ttu-id="888f1-223">(詳細については、「[Office アドインによって使用されるブラウザー](../concepts/browsers-used-by-office-web-add-ins.md)」を参照してください。)</span><span class="sxs-lookup"><span data-stu-id="888f1-223">(For more information, see [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).)</span></span>
    * <span data-ttu-id="888f1-224">`Office.initialize` メソッドへの割り当てが、`getGraphAccessTokenButton` ボタン クリック イベントへのハンドラーの割り当てになります。</span><span class="sxs-lookup"><span data-stu-id="888f1-224">An assignment to the `Office.initialize` method that, in turn, assigns a handler to the `getGraphAccessTokenButton` button click event.</span></span>
    * <span data-ttu-id="888f1-225">`showResult` メソッドは、作業ウィンドウの下側に Microsoft Graph から返されたデータ (またはエラー メッセージ) を表示するものです。</span><span class="sxs-lookup"><span data-stu-id="888f1-225">A `showResult` method that will display data returned from Microsoft Graph (or an error message) at the bottom of the task pane.</span></span>
    * <span data-ttu-id="888f1-226">`logErrors` メソッドは、エンド ユーザーを対象としていないエラーをコンソールにログ出力するものです。</span><span class="sxs-lookup"><span data-stu-id="888f1-226">A `logErrors` method that will log to console errors that are not intended for the end user.</span></span>
    * <span data-ttu-id="888f1-227">SSO がサポートされていない、または SSO がエラーになっているシナリオでアドインが使用するフォールバック認証システムを実装するコード。</span><span class="sxs-lookup"><span data-stu-id="888f1-227">Code that implements the fallback authorization system that the add-in will use in scenarios where SSO is not supported or has errored.</span></span>

1. <span data-ttu-id="888f1-228">`Office.initialize` への割り当ての下に、次に示すコードを追加します。</span><span class="sxs-lookup"><span data-stu-id="888f1-228">Below the assignment to `Office.initialize`, add the code below.</span></span> <span data-ttu-id="888f1-229">このコードについては、次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="888f1-229">Note the following about this code:</span></span>

    * <span data-ttu-id="888f1-230">アドインのエラー処理により、アクセス トークンの取得が別のオプションのセットを使用して自動的に再試行されることがあります。</span><span class="sxs-lookup"><span data-stu-id="888f1-230">The error-handling in the add-in will sometimes automatically attempt a second time to get an access token, using a different set of options.</span></span> <span data-ttu-id="888f1-231">カウンター変数 `retryGetAccessToken` は、ユーザーがトークンを取得しようとしたときに繰り返し再試行されないように使用されます。</span><span class="sxs-lookup"><span data-stu-id="888f1-231">The counter variable `retryGetAccessToken` is used to ensure that the user isn't cycled repeatedly through failed attempts to get a token.</span></span>
    * <span data-ttu-id="888f1-232">`getGraphData` 関数は、ES6 `async` キーワードで定義されます。</span><span class="sxs-lookup"><span data-stu-id="888f1-232">The `getGraphData` function is defined with the ES6 `async` keyword.</span></span> <span data-ttu-id="888f1-233">ES6 構文を使用すると、Office アドインの SSO API の使用が非常に簡単になります。</span><span class="sxs-lookup"><span data-stu-id="888f1-233">Using ES6 syntax makes the SSO API in Office Add-ins much easier to to use.</span></span> <span data-ttu-id="888f1-234">これは、ソリューション内の、Internet Explorer でサポートされていない構文を使用する唯一のファイルです。</span><span class="sxs-lookup"><span data-stu-id="888f1-234">This is the only file in the solution that will use syntax that is not supported by Internet Explorer.</span></span> <span data-ttu-id="888f1-235">ファイル名に「ES6」というリマインダーが設定されています。</span><span class="sxs-lookup"><span data-stu-id="888f1-235">We put 'ES6' in the filename as a reminder.</span></span> <span data-ttu-id="888f1-236">このソリューションでは、tsc トランスパイラーを使用してこのファイルを ES5 にトランスパイルします。これにより、Office が UI に Internet Explorer を使用しているときにアドインが実行されます。</span><span class="sxs-lookup"><span data-stu-id="888f1-236">The solution uses the tsc transpiler to transpile this file to ES5, so that the add-in can run when Office is using Internet Explorer for the UI.</span></span> <span data-ttu-id="888f1-237">(プロジェクトのルートにある tsconfig.json ファイルを参照します。)</span><span class="sxs-lookup"><span data-stu-id="888f1-237">(See the tsconfig.json file in the root of the project.)</span></span>

    ```javascript
    var retryGetAccessToken = 0;

    async function getGraphData() {
        await getDataWithToken({ allowSignInPrompt: true, forMSGraphAccess: true });
    }
    ```

1. <span data-ttu-id="888f1-238">`getGraphData` 関数の下に、次の関数を追加します。</span><span class="sxs-lookup"><span data-stu-id="888f1-238">Below the `getGraphData` function add the following function.</span></span> <span data-ttu-id="888f1-239">後の手順で `handleClientSideErrors` 関数を作成することに注意してください。</span><span class="sxs-lookup"><span data-stu-id="888f1-239">Note that you create the `handleClientSideErrors` function in a later step.</span></span>

    ```javascript
    async function getDataWithToken() {
        try {

            // TODO 1: Get the bootstrap token and send it to the server to exchange
            //         for an access token to Microsoft Graphn and then get the data
            //         from Microsoft Graph.

        }
        catch (exception) {
            if (exception.code) {
                handleClientSideErrors(exception);
            }
            else {
                showResult(["EXCEPTION: " + JSON.stringify(exception)]);
            }
        }
    }
    ```

1. <span data-ttu-id="888f1-240">`TODO 1`を以下のように置き換えます。</span><span class="sxs-lookup"><span data-stu-id="888f1-240">Replace `TODO 1` with the following.</span></span> <span data-ttu-id="888f1-241">このコードについては、以下の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="888f1-241">About this code, note:</span></span>

    * <span data-ttu-id="888f1-242">`getAccessToken` は、Azure AD からブートストラップ トークンを取得し、アドインに戻るように Office に指示します。</span><span class="sxs-lookup"><span data-stu-id="888f1-242">`getAccessToken` tells Office to get a bootstrap token from Azure AD and return to the add-in.</span></span>
    * <span data-ttu-id="888f1-243">`allowSignInPrompt` は、ユーザーがまだ Office にサインインしていない場合、ユーザーにサインインするように求めるように Office に指示します。</span><span class="sxs-lookup"><span data-stu-id="888f1-243">`allowSignInPrompt` tells Office to prompt the user to sign in if the user isn't already signed into Office.</span></span>
    * <span data-ttu-id="888f1-244">`forMSGraphAccess` は、アドインが (ブートストラップ トークンをユーザー ID トークンとして使用するだけでなく) Microsoft Graph へのアクセス トークンのブートストラップ トークンを交換することを Office に通知します。</span><span class="sxs-lookup"><span data-stu-id="888f1-244">`forMSGraphAccess` tells Office that the add-in intends to swap the bootstrap token for an access token to Microsoft Graph (instead of just using the bootstrap token as a user ID token).</span></span> <span data-ttu-id="888f1-245">このオプションを設定すると、ユーザーのテナント管理者がアドインの同意を与えていない場合、Office はブートストラップ トークンの取得プロセスをキャンセルすることができます (そしてエラー コード 13012 が返されます)。</span><span class="sxs-lookup"><span data-stu-id="888f1-245">Setting this option gives Office a chance to cancel the process of getting a bootstrap token (and return error code 13012) if the user's tenant administrator has not granted consent to the add-in.</span></span> <span data-ttu-id="888f1-246">アドインのクライアント側コードが 13012 に返信するには、フォールバック認証システムに分岐します。</span><span class="sxs-lookup"><span data-stu-id="888f1-246">The add-in's client-side code can respond to the 13012 by branching to a fallback authorization system.</span></span> <span data-ttu-id="888f1-247">`forMSGraphAccess` が使用されず、管理者が同意を与えていない場合は、ブートストラップ トークンが返されますが、on-behalf-of フローと交換しようとするとエラーになります。</span><span class="sxs-lookup"><span data-stu-id="888f1-247">If the `forMSGraphAccess` is not used, and the admin has not granted consent, the bootstrap token is returned, but the attempt to exhange it with the on-behalf-of flow would result in an error.</span></span> <span data-ttu-id="888f1-248">したがって、`forMSGraphAccess` オプションを使用すると、アドインがフォールバック システムにすばやく分岐できます。</span><span class="sxs-lookup"><span data-stu-id="888f1-248">Thus, the `forMSGraphAccess` option enables the add-in to branch to the fallback system quickly.</span></span>
    * <span data-ttu-id="888f1-249">後の手順で `getData` 関数を作成します。</span><span class="sxs-lookup"><span data-stu-id="888f1-249">You create the `getData` function in a later step.</span></span>
    * <span data-ttu-id="888f1-250">`/api/values` パラメーターは、トークンを交換したり、Microsoft Graph を呼び出すためにアクセス トークンを使用したりする、サーバー側コントローラーの URL です。</span><span class="sxs-lookup"><span data-stu-id="888f1-250">The `/api/values` parameter is the URL of a server-side controller that will make the token exchange and use the access token it gets back to make the call to Microsoft Graph.</span></span>

    ```javascript
    let bootstrapToken = await OfficeRuntime.auth.getAccessToken({
        allowSignInPrompt: true,
        forMSGraphAccess: true });

    getData("/api/values", bootstrapToken);
    ```

1. <span data-ttu-id="888f1-251">`getGraphData` 関数の下に、次を追加します。</span><span class="sxs-lookup"><span data-stu-id="888f1-251">Below the `getGraphData` function, add the following.</span></span> <span data-ttu-id="888f1-252">このコードについては、以下の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="888f1-252">About this code, note:</span></span>

    * <span data-ttu-id="888f1-253">これは、SSO 認証システムおよびフォールバック認証システムの両方で使用されます。</span><span class="sxs-lookup"><span data-stu-id="888f1-253">It is used by both the SSO and the fallback authorization systems.</span></span>
    * <span data-ttu-id="888f1-254">`relativeUrl` パラメーターは、サーバー側のコントローラーです。</span><span class="sxs-lookup"><span data-stu-id="888f1-254">The `relativeUrl` parameter is a server-side controller.</span></span>
    * <span data-ttu-id="888f1-255">`accessToken` パラメーターは、ブートストラップ トークンまたはフル アクセス トークンにすることができます。</span><span class="sxs-lookup"><span data-stu-id="888f1-255">The `accessToken` parameter can be a bootstrap token or a full access token.</span></span>
    * <span data-ttu-id="888f1-256">`writeFileNamesToOfficeDocument` は、既にプロジェクトの一部です。</span><span class="sxs-lookup"><span data-stu-id="888f1-256">The `writeFileNamesToOfficeDocument` is already part of the project.</span></span>
    * <span data-ttu-id="888f1-257">後の手順で `handleServerSideErrors` 関数を作成します。</span><span class="sxs-lookup"><span data-stu-id="888f1-257">You create the `handleServerSideErrors` function in a later step.</span></span>

    ```javascript
    function getData(relativeUrl, accessToken) {

        $.ajax({
            url: relativeUrl,
            headers: { "Authorization": "Bearer " + accessToken },
            type: "GET"
        })
            .done(function (result) {
                writeFileNamesToOfficeDocument(result)
                    .then(function () {
                        showResult(["Your data has been added to the document."]);
                    })
                    .catch(function (error) {
                        showResult([JSON.stringify(error)]);
                    });
            })
            .fail(function (result) {
                handleServerSideErrors(result);
            });
    }
    ```

### <a name="handle-client-side-errors"></a><span data-ttu-id="888f1-258">クライアント側のエラーを処理する</span><span class="sxs-lookup"><span data-stu-id="888f1-258">Handle client-side errors</span></span>

1. <span data-ttu-id="888f1-259">`getData` 関数の下に、次の関数を追加します。</span><span class="sxs-lookup"><span data-stu-id="888f1-259">Below the `getData` function, add the following function.</span></span> <span data-ttu-id="888f1-260">`error.code`は数値であり、通常は 13xxx の範囲にあることを注意してください。</span><span class="sxs-lookup"><span data-stu-id="888f1-260">Note that `error.code` is a number, usually in the range 13xxx.</span></span>

    ```javascript
    function handleClientSideErrors(error) {
        switch (error.code) {

            // TODO 2: Handle errors where the add-in should NOT invoke
            //         the alternative system of authorization.

            // TODO 3: Handle errors where the add-in should invoke
            //         the alternative system of authorization.

        }
    }
    ```

1. <span data-ttu-id="888f1-261">`TODO 2`を以下のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="888f1-261">Replace `TODO 2` with the following code.</span></span> <span data-ttu-id="888f1-262">これらのエラーの詳細については、「[Office アドインの SSO のトラブルシューティング (Troubleshoot SSO in Office Add-ins)](troubleshoot-sso-in-office-add-ins.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="888f1-262">For more information about these errors, see [Troubleshoot SSO in Office Add-ins](troubleshoot-sso-in-office-add-ins.md).</span></span>

    ```javascript
    case 13001:
        // No one is signed into Office. If the add-in cannot be effectively used when no one
        // is logged into Office, then the first call of getAccessToken should pass the
        // `allowSignInPrompt: true` option.
        showResult(["No one is signed into Office. But you can use many of the add-ins functions anyway. If you want to log in, press the Get OneDrive File Names button again."]);
        break;
    case 13002:
        // The user aborted the consent prompt. If the add-in cannot be effectively used when consent
        // has not been granted, then the first call of getAccessToken should pass the `allowConsentPrompt: true` option.
        showResult(["You can use many of the add-ins functions even though you have not granted consent. If you want to grant consent, press the Get OneDrive File Names button again."]);
        break;
    case 13006:
        // Only seen in Office on the Web.
        showResult(["Office on the Web is experiencing a problem. Please sign out of Office, close the browser, and then start again."]);
        break;
    case 13008:
        // Only seen in Office on the Web.
        showResult(["Office is still working on the last operation. When it completes, try this operation again."]);
        break;
    case 13010:
        // Only seen in Office on the Web.
        showResult(["Follow the instructions to change your browser's zone configuration."]);
        break;
    ```

1. <span data-ttu-id="888f1-263">`TODO 3`を以下のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="888f1-263">Replace `TODO 3` with the following code.</span></span> <span data-ttu-id="888f1-264">その他のエラーが発生した場合、アドインはフォールバック認証システムに分岐します。</span><span class="sxs-lookup"><span data-stu-id="888f1-264">For all other errors, the add-in branches to the fallback authorization system.</span></span> <span data-ttu-id="888f1-265">これらのエラーの詳細については、「[Office アドインの SSO のトラブルシューティング](troubleshoot-sso-in-office-add-ins.md)」を参照してください。このアドインでは、フォールバック システムはユーザーが既にログインしている場合でもユーザーのサインインを要求するダイアログを開き、msal.js および暗黙的フローを使用して Microsoft Graph へのアクセス トークンを取得します。</span><span class="sxs-lookup"><span data-stu-id="888f1-265">For more information about these errors, see [Troubleshoot SSO in Office Add-ins](troubleshoot-sso-in-office-add-ins.md). In this add-in, the fallback system opens a dialog which requires the user to sign in, even if the user already is, and uses msal.js and the Implicit Flow to get an access token to Microsoft Graph.</span></span>

    ```javascript
    default:
        dialogFallback();
        break;
    ```

### <a name="handle-server-side-errors"></a><span data-ttu-id="888f1-266">サーバー側のエラーを処理する</span><span class="sxs-lookup"><span data-stu-id="888f1-266">Handle server-side errors</span></span>

1. <span data-ttu-id="888f1-267">`handleClientSideErrors` 関数の下に、次の関数を追加します。</span><span class="sxs-lookup"><span data-stu-id="888f1-267">Below the `handleClientSideErrors` function, add the following function.</span></span>

    ```javascript
    function handleServerSideErrors(result) {

    // TODO 4: Parse the JSON response.

    // TODO 5: Handle case where Microsoft Graph requires an additional form
    //         of authentication.

    // TODO 6: Handle other Azure AD errors

    }
    ```

1. <span data-ttu-id="888f1-268">`TODO 4` を以下のように置き換えます。</span><span class="sxs-lookup"><span data-stu-id="888f1-268">Replace `TODO 4` with the following.</span></span> <span data-ttu-id="888f1-269">このコードについては、MFA などが存在する前に ASP.NET エラー クラスが作成されたことに注意してください。</span><span class="sxs-lookup"><span data-stu-id="888f1-269">About this code, note that ASP.NET error classes were created before there was such a thing as MFA.</span></span> <span data-ttu-id="888f1-270">第 2 認証要素に対する要求をサーバー側の論理が処理する方法の副作用として、クライアントに送信されるサーバー側のエラーは **Message** プロパティがありますが、**ExceptionMessage** プロパティはありません。</span><span class="sxs-lookup"><span data-stu-id="888f1-270">As a side-effect of how our server-side logic handles the requests for a second authentication factor, the server-side error sent to the client has a **Message** property but no **ExceptionMessage** property.</span></span> <span data-ttu-id="888f1-271">ただし、他のすべてのエラーには **ExceptionMessage** プロパティがあるため、クライアント側のコードは両方の応答を解析する必要があります。</span><span class="sxs-lookup"><span data-stu-id="888f1-271">But all other errors will have a **ExceptionMessage** property, so the client-side code has to parse the response for both.</span></span> <span data-ttu-id="888f1-272">どちらか一方の変数が未定義になります。</span><span class="sxs-lookup"><span data-stu-id="888f1-272">Either one or the other variable will be undefined.</span></span>

    ```javascript
    var message = JSON.parse(result.responseText).Message;
    var exceptionMessage = JSON.parse(result.responseText).ExceptionMessage;
    ```

1. <span data-ttu-id="888f1-273">`TODO 5` を以下のように置き換えます。</span><span class="sxs-lookup"><span data-stu-id="888f1-273">Replace `TODO 5` with the following.</span></span> <span data-ttu-id="888f1-274">Microsoft Graph が認証の追加形式を必要とする場合、エラー AADSTS50076 が送信されます。</span><span class="sxs-lookup"><span data-stu-id="888f1-274">When Microsoft Graph requires an additional form of authentication, it sends error AADSTS50076.</span></span> <span data-ttu-id="888f1-275">これには、**Message.Claims** プロパティの追加要件に関する情報が含まれます。</span><span class="sxs-lookup"><span data-stu-id="888f1-275">It includes information about the additional requirement in the **Message.Claims** property.</span></span> <span data-ttu-id="888f1-276">これを処理するために、コードはブートストラップ トークンの取得を 2 回試行しますが、今回は `authChallenge` オプションの値として追加要素の要求が含まれます。これにより、Azure AD は、必要なすべての形式の認証をユーザーに要求します。</span><span class="sxs-lookup"><span data-stu-id="888f1-276">To handle this, the code makes a second attempt to get the bootstrap token, but this time it includes the request for an additional factor as the value of the `authChallenge` option, which tells Azure AD to prompt the user for all required forms of authentication.</span></span>

    ```javascript
    if (message) {
        if (message.indexOf("AADSTS50076") !== -1) {
            var claims = JSON.parse(message).Claims;
            var claimsAsString = JSON.stringify(claims);
            getDataWithToken({ authChallenge: claimsAsString });
            return;
        }
    }
    ```

1. <span data-ttu-id="888f1-277">`TODO 6` を以下のように置き換えます。</span><span class="sxs-lookup"><span data-stu-id="888f1-277">Replace `TODO 6` with the following.</span></span>

    ```javascript
    if (exceptionMessage) {

        // TODO 7: Handle case where bootstrap token has expired.

        // TODO 8: Handle all other Azure AD errors.
    }
    ```

1. <span data-ttu-id="888f1-278">`TODO 7` を以下のように置き換えます。</span><span class="sxs-lookup"><span data-stu-id="888f1-278">Replace `TODO 7` with the following.</span></span> <span data-ttu-id="888f1-279">まれにブートストラップ トークンが Office の検証時に期限切れにならず、交換のために Azure AD に送信されるまでの間に期限切れになることがあることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="888f1-279">Note that on rare occasions the bootstrap token is unexpired when Office validates it, but expires by the time it is sent to Azure AD for exchange.</span></span> <span data-ttu-id="888f1-280">Azure AD はエラー AADSTS500133 で応答します。</span><span class="sxs-lookup"><span data-stu-id="888f1-280">Azure AD will respond with error AADSTS500133.</span></span> <span data-ttu-id="888f1-281">この場合、コードは SSO API を呼び戻します (ただし、1 回のみ)。</span><span class="sxs-lookup"><span data-stu-id="888f1-281">When this happens, the code  recalls the SSO API (but no more than once).</span></span> <span data-ttu-id="888f1-282">今回は、Office が期限切れになっていない新しいブートストラップ トークンを返します。</span><span class="sxs-lookup"><span data-stu-id="888f1-282">This time Office returns a new unexpired bootstrap token.</span></span>

    ```javascript
    if ((exceptionMessage.indexOf("AADSTS500133") !== -1)
        && (retryGetAccessToken <= 0)) {

        retryGetAccessToken++;
        getGraphData();
    }
    ```

1. <span data-ttu-id="888f1-283">`TODO 8` を以下のように置き換えます。</span><span class="sxs-lookup"><span data-stu-id="888f1-283">Replace `TODO 8` with the following.</span></span>

    ```javascript
    else {
        dialogFallback();
    }
    ```

1. <span data-ttu-id="888f1-284">ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="888f1-284">Save the file.</span></span>

## <a name="code-the-server-side"></a><span data-ttu-id="888f1-285">サーバー側のコードを作成する</span><span class="sxs-lookup"><span data-stu-id="888f1-285">Code the server side</span></span>

### <a name="configure-the-owin-middleware"></a><span data-ttu-id="888f1-286">OWIN ミドルウェアを構成する</span><span class="sxs-lookup"><span data-stu-id="888f1-286">Configure the OWIN middleware</span></span>

1. <span data-ttu-id="888f1-287">**Office-Add-in-ASPNET-SSO-WebAPI** プロジェクトのルートにある Startup.cs ファイルを開き、**スタートアップ** クラスに次のメソッドを追加します。</span><span class="sxs-lookup"><span data-stu-id="888f1-287">Open the Startup.cs file in the root of the **Office-Add-in-ASPNET-SSO-WebAPI** project and add the following method to the **Startup** class.</span></span> <span data-ttu-id="888f1-288">`ConfigureAuth` メソッドは、この後の手順で作成することに注意してください。</span><span class="sxs-lookup"><span data-stu-id="888f1-288">Note that you create the `ConfigureAuth` method in a later step.</span></span>

    ```csharp
    public void Configuration(IAppBuilder app)
    {
        ConfigureAuth(app);
    }
    ```

1. <span data-ttu-id="888f1-289">ファイルを保存して閉じます。</span><span class="sxs-lookup"><span data-stu-id="888f1-289">Save and close the file.</span></span>

1. <span data-ttu-id="888f1-290">**App_Start** フォルダーを右クリックして、**[追加] > [クラス]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="888f1-290">Right-click the **App_Start** folder and select **Add > Class**.</span></span>

1. <span data-ttu-id="888f1-291">**[新しい項目の追加]** ダイアログで、ファイルに「**Startup.Auth.cs**」という名前を付けて **[追加]** をクリックします。</span><span class="sxs-lookup"><span data-stu-id="888f1-291">In the **Add new item** dialog name the file **Startup.Auth.cs** and then click **Add**.</span></span>

1. <span data-ttu-id="888f1-292">新しいファイルで名前空間の名前を `Office_Add_in_ASPNET_SSO_WebAPI` に短縮します。</span><span class="sxs-lookup"><span data-stu-id="888f1-292">Shorten the namespace name in the new file to `Office_Add_in_ASPNET_SSO_WebAPI`.</span></span>

1. <span data-ttu-id="888f1-293">ファイルの先頭に、次に示す `using` ステートメントがすべて揃っていることを確認します。</span><span class="sxs-lookup"><span data-stu-id="888f1-293">Ensure that all of the following `using` statements are at the top of the file.</span></span>

    ```csharp
    using Owin;
    using Microsoft.IdentityModel.Tokens;
    using System.Configuration;
    using Microsoft.Owin.Security.OAuth;
    using Microsoft.Owin.Security.Jwt;
    using Office_Add_in_ASPNET_SSO_WebAPI.App_Start;
    ```

1. <span data-ttu-id="888f1-p148">`Startup` クラスの宣言にキーワード `partial` を追加します (まだ追加されていない場合)。これは、次のようになります。</span><span class="sxs-lookup"><span data-stu-id="888f1-p148">Add the keyword `partial` to the declaration of the `Startup` class, if it is not already there. It should look like this:</span></span>

    `public partial class Startup`

1. <span data-ttu-id="888f1-p149">次に示すメソッドを `Startup` クラスに追加します。このメソッドでは、クライアント側の Home.js ファイルの `getData` メソッドから渡されたアクセス トークンを OWIN ミドルウェアで検証する方法を指定します。承認プロセスは、`[Authorize]` 属性で修飾された Web API エンドポイントが呼び出されたときには必ずトリガーされます。</span><span class="sxs-lookup"><span data-stu-id="888f1-p149">Add the following method to the `Startup` class. This method specifies how the OWIN middleware will validate the access tokens that are passed to it from the `getData` method in the client-side Home.js file. The authorization process is triggered whenever a Web API endpoint that is decorated with the `[Authorize]` attribute is called.</span></span>

    ```csharp
    public void ConfigureAuth(IAppBuilder app)
    {
        // TODO 1: Configure the validation settings

        // TODO 2: Specify the type of authorization and the discovery endpoint
        //        of the secure token service.
    }
    ```

1. <span data-ttu-id="888f1-299">`TODO 1` を以下のように置き換えます。</span><span class="sxs-lookup"><span data-stu-id="888f1-299">Replace the `TODO 1` with the following.</span></span> <span data-ttu-id="888f1-300">このコードの注意点は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="888f1-300">Note about this code:</span></span>

    * <span data-ttu-id="888f1-301">このコードでは、Office ホストから得られるブートストラップ トークンで指定された対象ユーザーが web.config で指定された値と一致する必要があることを OWIN に指示します。</span><span class="sxs-lookup"><span data-stu-id="888f1-301">The code instructs OWIN to ensure that the audience specified in the bootstrap token that comes from the Office host must match the value specified in the web.config.</span></span>
    * <span data-ttu-id="888f1-302">Microsoft アカウントには、組織のテナント GUID とは異なる発行者 GUID があります。そのため、両方の種類のアカウントをサポートするために、発行者は検証されません。</span><span class="sxs-lookup"><span data-stu-id="888f1-302">Microsoft Accounts have an issuer GUID that is different from any organizational tenant GUID, so to support both kinds of accounts, we do not validate the issuer.</span></span>
    * <span data-ttu-id="888f1-303">`SaveSigninToken` を `true` に設定することで、OWIN は Office ホストからの生のブートストラップ トークンを保存するようになります。</span><span class="sxs-lookup"><span data-stu-id="888f1-303">Setting `SaveSigninToken` to `true` causes OWIN to save the raw bootstrap token from the Office host.</span></span> <span data-ttu-id="888f1-304">これは、アドインが代理フローで Microsoft Graph へのアクセス トークンを取得するために必要になります。</span><span class="sxs-lookup"><span data-stu-id="888f1-304">The add-in needs it to obtain an access token to Microsoft Graph with the on-behalf-of flow.</span></span>
    * <span data-ttu-id="888f1-305">OWIN ミドルウェアでは、スコープは検証されません。</span><span class="sxs-lookup"><span data-stu-id="888f1-305">Scopes are not validated by the OWIN middleware.</span></span> <span data-ttu-id="888f1-306">`access_as_user` が含まれている必要があるブートストラップ トークンのスコープは、コントローラーで検証されます。</span><span class="sxs-lookup"><span data-stu-id="888f1-306">The scopes of the bootstrap token, which should include `access_as_user`, is validated in the controller.</span></span>

    ```csharp
    TokenValidationParameters tvps = new TokenValidationParameters
    {
        ValidAudience = ConfigurationManager.AppSettings["ida:Audience"],
        ValidateIssuer = false,
        SaveSigninToken = true
    };
    ```

1. <span data-ttu-id="888f1-307">`TODO 2`を以下のように置き換えます。</span><span class="sxs-lookup"><span data-stu-id="888f1-307">Replace `TODO 2` with the following.</span></span> <span data-ttu-id="888f1-308">このコードの注意点は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="888f1-308">Note about this code:</span></span>

    * <span data-ttu-id="888f1-309">より一般的な `UseWindowsAzureActiveDirectoryBearerAuthentication` は Azure AD V2 エンドポイントに準拠していないため、その代わりとしてメソッド `UseOAuthBearerAuthentication` が呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="888f1-309">The method `UseOAuthBearerAuthentication` is called instead of the more common `UseWindowsAzureActiveDirectoryBearerAuthentication` because the latter is not compatible with the Azure AD V2 endpoint.</span></span>
    * <span data-ttu-id="888f1-310">このメソッドに渡される URL は、Office ホストから受け取ったブートストラップ トークンの署名の検証に必要になるキーを取得するための方法を OWIN ミドルウェアが取得する場所になります。</span><span class="sxs-lookup"><span data-stu-id="888f1-310">The URL that is passed to the method is where the OWIN middleware obtains instructions for getting the key it needs to verify the signature on the bootstrap token received from the Office host.</span></span> <span data-ttu-id="888f1-311">URL の権威セグメントは、web.config から取得されます。これは「common」という文字列か、シングルテナント アドインの場合は GUID です。</span><span class="sxs-lookup"><span data-stu-id="888f1-311">The Authority segment of the URL comes from the web.config. It is either the string "common" or, for a single-tenant add-in, a GUID.</span></span>

    ```csharp
    string[] endAuthoritySegments = { "oauth2/v2.0" };
    string[] parsedAuthority = ConfigurationManager.AppSettings["ida:Authority"].Split(endAuthoritySegments, System.StringSplitOptions.None);
    string wellKnownURL = parsedAuthority[0] + "v2.0/.well-known/openid-configuration";

    app.UseOAuthBearerAuthentication(new OAuthBearerAuthenticationOptions
    {
        AccessTokenFormat = new JwtFormat(tvps, new OpenIdConnectCachingSecurityTokenProvider(wellKnownURL))
    });
    ```

1. <span data-ttu-id="888f1-312">ファイルを保存して閉じます。</span><span class="sxs-lookup"><span data-stu-id="888f1-312">Save and close the file.</span></span>

### <a name="create-the-apivalues-controller"></a><span data-ttu-id="888f1-313">/api/values コントローラーを作成する</span><span class="sxs-lookup"><span data-stu-id="888f1-313">Create the /api/values controller</span></span>

1. <span data-ttu-id="888f1-314">ファイル **Controllers\ValueController.cs** を開きます。</span><span class="sxs-lookup"><span data-stu-id="888f1-314">Open the file **Controllers\ValueController.cs**.</span></span> <span data-ttu-id="888f1-315">このコントローラーは、SSO システムがブートストラップ トークンを正常に取得した場合に使用されます。</span><span class="sxs-lookup"><span data-stu-id="888f1-315">This controller is used when the SSO system has successfully obtained a bootstrap token.</span></span> <span data-ttu-id="888f1-316">フォールバック認証システムの一部として使用されることはありません。</span><span class="sxs-lookup"><span data-stu-id="888f1-316">It is not used as part of the fallback authorization system.</span></span> <span data-ttu-id="888f1-317">そのシステムで AzureADAuthController が使用されました。これは、自動的に作成されます。</span><span class="sxs-lookup"><span data-stu-id="888f1-317">That system used the AzureADAuthController, which has been created for you.</span></span>

1. <span data-ttu-id="888f1-318">ファイルの先頭に、次に示す `using` ステートメントがあることを確認します。</span><span class="sxs-lookup"><span data-stu-id="888f1-318">Ensure that the following `using` statements are at the top of the file.</span></span>

    ```csharp
    using Microsoft.Identity.Client;
    using System.Configuration;
    using System.Linq;
    using System.Security.Claims;
    using System.Threading.Tasks;
    using System.Web.Http;
    using System;
    using System.Net;
    using System.Net.Http;
    using Office_Add_in_ASPNET_SSO_WebAPI.Helpers;
    ```

1. <span data-ttu-id="888f1-p156">`ValuesController` を宣言している行のすぐ上に、属性 `[Authorize]` を追加します。これにより、アドインはコントローラー メソッドが呼び出されたときに、最後の手順で構成した承認プロセスを必ず実行するようになります。アドインへの有効なアクセス トークンを持つ呼び出し元のみが、コントローラーのメソッドを起動できます。</span><span class="sxs-lookup"><span data-stu-id="888f1-p156">Just above the line that declares the `ValuesController`, add the `[Authorize]` attribute. This ensures that your add-in will run the authorization process that you configured in the last procedure whenever a controller method is called. Only callers with a valid access token to your add-in can invoke the methods of the controller.</span></span>

1. <span data-ttu-id="888f1-322">次のメソッドを `ValuesController` に追加します。</span><span class="sxs-lookup"><span data-stu-id="888f1-322">Add the following method to the `ValuesController`.</span></span> <span data-ttu-id="888f1-323">戻り値は、`Task<IEnumerable<string>>` ではなく `GET api/values` メソッドでより一般的な `Task<HttpResponseMessage>` になる点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="888f1-323">Note that the return value is `Task<HttpResponseMessage>` instead of `Task<IEnumerable<string>>` as would be more common for a `GET api/values` method.</span></span> <span data-ttu-id="888f1-324">これは、OAuth 認証論理が ASP.NET フィルターではなく、コントローラーに存在する必要があるということの副作用です。</span><span class="sxs-lookup"><span data-stu-id="888f1-324">This is a side effect of that fact that the OAuth  authorization logic must be in the controller, instead of in an ASP.NET filter.</span></span> <span data-ttu-id="888f1-325">その論理の一部のエラーの条件では、アドインのクライアントに HTTP 応答オブジェクトが送信される必要があります。</span><span class="sxs-lookup"><span data-stu-id="888f1-325">Some error conditions in that logic require that an HTTP Response object be sent to the add-in's client.</span></span>

    ```csharp
    // GET api/values
    public async Task<HttpResponseMessage> Get()
    {
        // TODO 1: Validate the scopes of the bootstrap token.

        // TODO 2: Assemble all the information that is needed to get a
        //        token for Microsoft Graph using the on-behalf-of flow.

        // TODO 3: Get the access token for Microsoft Graph.

        // TODO 4: Use the token to call Microsoft Graph.
    }
    ```

1. <span data-ttu-id="888f1-326">`TODO1` を次のコードに置き換えて、`access_as_user` を含むトークンで指定されているスコープを検証します。</span><span class="sxs-lookup"><span data-stu-id="888f1-326">Replace `TODO1` with the following code to validate that the scopes that are specified in the token include `access_as_user`.</span></span> <span data-ttu-id="888f1-327">`SendErrorToClient` メソッドの第 2 パラメーターは、**Exception** オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="888f1-327">Note that the second parameter of the `SendErrorToClient` method is an **Exception** object.</span></span> <span data-ttu-id="888f1-328">この場合、コードは `null` を渡します。これは、**Exception** オブジェクトが含まれていることで、生成される HTTP 応答には **Message** プロパティが含められなくなるためです。</span><span class="sxs-lookup"><span data-stu-id="888f1-328">In this case, the code passes `null` because including the **Exception** object blocks the inclusion of the **Message** property in the HTTP Response that is generated.</span></span>


    ```csharp
    string[] addinScopes = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/scope").Value.Split(' ');
    if (!(addinScopes.Contains("access_as_user")))
    {
        return HttpErrorHelper.SendErrorToClient(HttpStatusCode.Unauthorized, null, "Missing access_as_user.");
    }
    ```

1. <span data-ttu-id="888f1-329">`TODO 2` を次のコードに置き換えて、「代理」フローを使用して Microsoft Graph のトークンを取得するために必要なすべての情報を編成します。</span><span class="sxs-lookup"><span data-stu-id="888f1-329">Replace `TODO 2` with the following code to assemble all the information that is needed to get a token for Microsoft Graph using the "on behalf of" flow.</span></span> <span data-ttu-id="888f1-330">このコードについては、以下の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="888f1-330">About this code, note:</span></span>

    * <span data-ttu-id="888f1-p160">アドインは、Office ホストとユーザーがアクセスする必要のあるリソース (または対象ユーザー) の役割を果たさなくなります。この時点で、それ自体が Microsoft Graph にアクセスする必要があるクライアントになります。`ConfidentialClientApplication` は MSAL の「クライアント コンテキスト」オブジェクトになります。</span><span class="sxs-lookup"><span data-stu-id="888f1-p160">Your add-in is no longer playing the role of a resource (or audience) to which the Office host and user need access. Now it is itself a client that needs access to Microsoft Graph. `ConfidentialClientApplication` is the MSAL “client context” object.</span></span>
    * <span data-ttu-id="888f1-334">MSAL.NET 3.x.x からは、`bootstrapContext` は単なるブートストラップ トークンです。</span><span class="sxs-lookup"><span data-stu-id="888f1-334">Beginning with MSAL.NET 3.x.x, the `bootstrapContext` is just the bootstrap token itself.</span></span>
    * <span data-ttu-id="888f1-335">権威は、web.config から取得されます。これは「common」という文字列か、シングルテナント アドインの場合は GUID です。</span><span class="sxs-lookup"><span data-stu-id="888f1-335">The Authority comes from the web.config. It is either the string "common" or, for a single-tenant add-in, a GUID.</span></span>
    * <span data-ttu-id="888f1-p161">MSAL では `openid`、`offline_access` の各スコープが機能することが必要ですが、コードがこれらを重複して要求するとエラーがスローされます。 コードが `profile` を要求した場合にもエラーがスローされます。それは、実際には Office ホスト アプリケーションがアドインの Web アプリケーションに対しトークンを取得するときだけに使用します。 そのため、`Files.Read.All` のみが明示的に要求されます。</span><span class="sxs-lookup"><span data-stu-id="888f1-p161">MSAL requires the `openid` and `offline_access` scopes to function, but it throws an error if your code redundantly requests them. It will also throw an error if your code requests `profile`, which is really only used when the Office host application gets the token to your add-in's web application. So only `Files.Read.All` is explicitly requested.</span></span>

    ```csharp
    string bootstrapContext = ClaimsPrincipal.Current.Identities.First().BootstrapContext.ToString();
    UserAssertion userAssertion = new UserAssertion(bootstrapContext);

    var cca = ConfidentialClientApplicationBuilder.Create(ConfigurationManager.AppSettings["ida:ClientID"])
                                                    .WithRedirectUri("https://localhost:44355")
                                                    .WithClientSecret(ConfigurationManager.AppSettings["ida:Password"])
                                                    .WithAuthority(ConfigurationManager.AppSettings["ida:Authority"])
                                                    .Build();

    string[] graphScopes = { "https://graph.microsoft.com/Files.Read.All" };
    ```

1. <span data-ttu-id="888f1-p162">`TODO 3` を次のコードに置き換えます。このコードの注意点は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="888f1-p162">Replace `TODO 3` with the following code. Note about this code:</span></span>

    * <span data-ttu-id="888f1-341">`ConfidentialClientApplication.AcquireTokenOnBehalfOfAsync` メソッドは、最初にメモリ内の MSAL キャッシュで一致するアクセス トークンを探します。</span><span class="sxs-lookup"><span data-stu-id="888f1-341">The `ConfidentialClientApplication.AcquireTokenOnBehalfOfAsync` method will first look in the MSAL cache, which is in memory, for a matching access token.</span></span> <span data-ttu-id="888f1-342">それが見つからなかった場合にのみ、Azure AD V2 エンドポイントで代理フローを開始します。</span><span class="sxs-lookup"><span data-stu-id="888f1-342">Only if there isn't one, does it initiate the on-behalf-of flow with the Azure AD V2 endpoint.</span></span>
    * <span data-ttu-id="888f1-343">`MsalServiceException` 以外の種類の例外は、意図的にキャッチしていないため、`500 Server Error` メッセージとしてクライアントに伝達されます。</span><span class="sxs-lookup"><span data-stu-id="888f1-343">Any exceptions that are not of type `MsalServiceException` are intentionally not caught, so they will propagate to the client as `500 Server Error` messages.</span></span>

    ```csharp
    AcquireTokenOnBehalfOfParameterBuilder parameterBuilder = null;
    AuthenticationResult authResult = null;
    try
    {
        parameterBuilder = cca.AcquireTokenOnBehalfOf(graphScopes, userAssertion);
        authResult = await parameterBuilder.ExecuteAsync();
    }
    catch (MsalServiceException e)
    {
        // TODO 3a: Handle request for multi-factor authentication.

        // TODO 3b: Handle lack of consent and invalid scope (permission).

        // TODO 3c: Handle all other MsalServiceExceptions.
    }
    ```

1. <span data-ttu-id="888f1-344">`TODO 3a`を以下のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="888f1-344">Replace `TODO 3a` with the following code.</span></span> <span data-ttu-id="888f1-345">このコードについては、以下の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="888f1-345">About this code, note:</span></span>

    * <span data-ttu-id="888f1-346">Microsoft Graph リソースが多要素認証を必要としているときに、その認証をユーザーがまだ指定していない場合、Azure AD はエラー `AADSTS50076` と **Claims** プロパティを含む「400 要求が正しくありません」を返します。</span><span class="sxs-lookup"><span data-stu-id="888f1-346">If multi-factor authentication is required by the Microsoft Graph resource and the user has not yet provided it, Azure AD will return "400 Bad Request" with error `AADSTS50076` and a **Claims** property.</span></span> <span data-ttu-id="888f1-347">MSAL は、この情報と共に **MsalUiRequiredException** (**MsalServiceException** から継承) をスローします。</span><span class="sxs-lookup"><span data-stu-id="888f1-347">MSAL throws a **MsalUiRequiredException** (which inherits from **MsalServiceException**) with this information.</span></span>
    * <span data-ttu-id="888f1-348">**Claims** プロパティの値は、クライアントに渡す必要があります。クライアントは、その値を Office ホストに渡します。Office ホストは、その値を新しいブートストラップ トークンの要求に含めます。</span><span class="sxs-lookup"><span data-stu-id="888f1-348">The **Claims** property value must be passed to the client which should pass it to the Office host, which then includes it in a request for a new bootstrap token.</span></span> <span data-ttu-id="888f1-349">Azure AD は、認証のすべての要求されたフォームをユーザーに示します。</span><span class="sxs-lookup"><span data-stu-id="888f1-349">Azure AD will prompt the user for all required forms of authentication.</span></span>
    * <span data-ttu-id="888f1-p167">例外から HTTP 応答を作成する API は、**Claims** プロパティを認識しないため、このプロパティを応答オブジェクトに含めません。 これが含まれたメッセージを手動で作成する必要があります。 ただし、カスタムの **Message** プロパティは **ExceptionMessage** プロパティの作成を妨げるため、クライアントがエラー ID `AADSTS50076` を取得するには、その ID をカスタムの **Message** に追加する以外に方法はありません。 クライアントの JavaScript では、応答に **Message** または **ExceptionMessage** が含まれているかどうかを検出する必要があるため、どちらを読み取るかを認識します。</span><span class="sxs-lookup"><span data-stu-id="888f1-p167">The APIs that create HTTP Responses from exceptions don't know about the **Claims** property, so they don't include it in the response object. We have to manually create a message that includes it. A custom **Message** property, however, blocks the creation of an **ExceptionMessage** property, so the only way to get the error ID `AADSTS50076` to the client is to add it to the custom **Message**. JavaScript in the client will need to discover if a response has a **Message** or **ExceptionMessage**, so it knows which to read.</span></span>
    * <span data-ttu-id="888f1-354">カスタム メッセージは、JSON として書式設定されているため、クライアント側の JavaScript は既知の JavaScript `JSON` オブジェクトのメソッドでメッセージを解析できます。</span><span class="sxs-lookup"><span data-stu-id="888f1-354">The custom message is formatted as JSON so that the client-side JavaScript can parse it with well-known JavaScript `JSON` object methods.</span></span>

    ```csharp
    if (e.Message.StartsWith("AADSTS50076"))
    {
        string responseMessage = String.Format("{{\"AADError\":\"AADSTS50076\",\"Claims\":{0}}}", e.Claims);
        return HttpErrorHelper.SendErrorToClient(HttpStatusCode.Forbidden, null, responseMessage);
    }
    ```

1. <span data-ttu-id="888f1-355">`TODO 3b`を以下のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="888f1-355">Replace `TODO 3b` with the following code.</span></span> <span data-ttu-id="888f1-356">このコードについては、以下の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="888f1-356">About this code, note:</span></span>

    * <span data-ttu-id="888f1-357">Azure AD の呼び出しにユーザーまたはテナント管理者のどちらも同意していない (または同意が取り消された) スコープ (アクセス許可) が少なくとも 1 つ含まれていると、Azure AD はエラー `AADSTS65001` と共に「400 要求が正しくありません」を返します。</span><span class="sxs-lookup"><span data-stu-id="888f1-357">If the call to Azure AD contained at least one scope (permission) for which neither the user nor a tenant administrator has consented (or consent was revoked), Azure AD will return "400 Bad Request" with error `AADSTS65001`.</span></span> <span data-ttu-id="888f1-358">MSAL は、この情報と共に **MsalUiRequiredException** をスローします。</span><span class="sxs-lookup"><span data-stu-id="888f1-358">MSAL throws a **MsalUiRequiredException** with this information.</span></span>
    *  <span data-ttu-id="888f1-359">Azure AD の呼び出しに Azure AD が認識しないスコープが少なくとも 1 つ含まれていると、AAD はエラー `AADSTS70011` と共に「400 要求が正しくありません」を返します。</span><span class="sxs-lookup"><span data-stu-id="888f1-359">If the call to Azure AD contained at least one scope that Azure AD does not recognize, AAD returns "400 Bad Request" with error `AADSTS70011`.</span></span> <span data-ttu-id="888f1-360">MSAL は、この情報と共に **MsalUiRequiredException** をスローします。</span><span class="sxs-lookup"><span data-stu-id="888f1-360">MSAL throws a **MsalUiRequiredException** with this information.</span></span>
    *  <span data-ttu-id="888f1-361">すべての説明が含まれている理由は、別の条件で 70011 が返されたときに、このアドインでは無効なスコープの存在を意味する場合のみを処理する必要があるためです。</span><span class="sxs-lookup"><span data-stu-id="888f1-361">The entire description is included because 70011 is returned in other conditions and it should only be handled in this add-in when it means that there is an invalid scope.</span></span>
    *  <span data-ttu-id="888f1-p171">**MsalUiRequiredException** オブジェクトが `SendErrorToClient` に渡されます。これにより、エラー情報を格納している **ExceptionMessage** プロパティが HTTP 応答に含まれるようにします。</span><span class="sxs-lookup"><span data-stu-id="888f1-p171">The **MsalUiRequiredException** object is passed to `SendErrorToClient`. This ensures that an **ExceptionMessage** property that contains the error information is included in the HTTP Response.</span></span>

    ```csharp
    if ((e.Message.StartsWith("AADSTS65001")) || (e.Message.StartsWith("AADSTS70011: The provided value for the input parameter 'scope' is not valid.")))
    {
        return HttpErrorHelper.SendErrorToClient(HttpStatusCode.Forbidden, e, null);
    }
    ```

1. <span data-ttu-id="888f1-364">`TODO 3c` を次のコードに置き換えて、他のすべての **MsalServiceException** を処理します。</span><span class="sxs-lookup"><span data-stu-id="888f1-364">Replace `TODO 3c` with the following code to handle all other **MsalServiceException**s.</span></span> <span data-ttu-id="888f1-365">前に説明したように、</span><span class="sxs-lookup"><span data-stu-id="888f1-365">As noted earlier,</span></span>

    ```csharp
    else
    {
        throw e;
    }
    ```

1. <span data-ttu-id="888f1-366">`TODO 4` を次のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="888f1-366">Replace `TODO 4` with the following code.</span></span> <span data-ttu-id="888f1-367">`GraphApiHelper.GetOneDriveFileNames` メソッドは、自動的に作成されます。これは、Microsoft Graph にデータを要求し、アクセス トークンを含めます。</span><span class="sxs-lookup"><span data-stu-id="888f1-367">The `GraphApiHelper.GetOneDriveFileNames` method, which has been created for you, makes the request for data to Microsoft Graph and includes the access token.</span></span>

    ```csharp
    return await GraphApiHelper.GetOneDriveFileNames(authResult.AccessToken);
    ```

1. <span data-ttu-id="888f1-368">ファイルを保存して閉じます。</span><span class="sxs-lookup"><span data-stu-id="888f1-368">Save and close the file.</span></span>

## <a name="run-the-solution"></a><span data-ttu-id="888f1-369">ソリューションを実行する</span><span class="sxs-lookup"><span data-stu-id="888f1-369">Run the solution</span></span>

1. <span data-ttu-id="888f1-370">Visual Studio ソリューション ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="888f1-370">Open the Visual Studio solution file.</span></span>
1. <span data-ttu-id="888f1-371">[**ビルド**] メニューで [**ソリューションのクリーン**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="888f1-371">On the **Build** menu, select **Clean Solution**.</span></span> <span data-ttu-id="888f1-372">終了したら、[**ビルド**] メニューをもう一度開き、[**ソリューションのビルド**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="888f1-372">When it finishes, open the **Build** menu again and select **Build Solution**.</span></span>
1. <span data-ttu-id="888f1-373">[**ソリューション エクスプローラー**] で、[**Office-Add-in-ASPNET-SSO**] を選択します (一番上のソリューション ノードではなく、「WebAPI」で終わる名前のプロジェクトではありません)。</span><span class="sxs-lookup"><span data-stu-id="888f1-373">In **Solution Explorer**, select the **Office-Add-in-ASPNET-SSO** project node (not the top solution node and not the project whose name ends in "WebAPI").</span></span>
1. <span data-ttu-id="888f1-374">[**プロパティ**] ウィンドウで、[**ドキュメントの開始**] ドロップダウンを開き、3 つのオプション (Excel、Word、または PowerPoint) のいずれかを選択します。</span><span class="sxs-lookup"><span data-stu-id="888f1-374">In the **Properties** pane, open the **Start Document** drop down and choose one of the three options (Excel, Word, or PowerPoint).</span></span>

    ![必要な Office ホスト アプリケーション (Excel、PowerPoint、または Word) を選択する](../images/SelectHost.JPG)

1. <span data-ttu-id="888f1-376">F5 キーを押します。</span><span class="sxs-lookup"><span data-stu-id="888f1-376">Press F5.</span></span>
1. <span data-ttu-id="888f1-377">Office アプリケーションの [**ホーム**] リボンで、[**SSO ASP.NET**] グループの [**アドインの表示**] を選択して、タスク ウィンドウ アドインを開きます。</span><span class="sxs-lookup"><span data-stu-id="888f1-377">In the Office application, on the **Home** ribbon, select the **Show Add-in** in the **SSO ASP.NET** group to open the task pane add-in.</span></span>
1. <span data-ttu-id="888f1-378">[**OneDrive ファイル名の取得**] ボタンをクリックします。</span><span class="sxs-lookup"><span data-stu-id="888f1-378">Click the **Get OneDrive File Names** button.</span></span> <span data-ttu-id="888f1-379">職場または学校の (Office 365) アカウントまたは Microsoft アカウントで Office にログインし、SSO が期待どおりに機能している場合、OneDrive for Business の最初の 10 個のファイル名およびフォルダー名がタスク ウィンドウに表示されます。</span><span class="sxs-lookup"><span data-stu-id="888f1-379">If you are logged into Office with either a Work or School (Office 365) account or Microsoft Account, and SSO is working as expected, the first 10 file and folder names in your OneDrive for Business are displayed on the task pane.</span></span> <span data-ttu-id="888f1-380">ログインしていない、または SSO をサポートしていないシナリオにいる場合、もしくは何らかの理由で SSO が機能していない場合には、ログインするように求められます。</span><span class="sxs-lookup"><span data-stu-id="888f1-380">If you are not logged in, or you are in a scenario that does not support SSO, or SSO is not working for any reason, you will be prompted to log in.</span></span> <span data-ttu-id="888f1-381">ログインすると、ファイル名およびフォルダー名が表示されます。</span><span class="sxs-lookup"><span data-stu-id="888f1-381">After you log in, the file and folder names appear.</span></span>
