---
title: シングル サインオンを使用する ASP.NET Office アドインを作成する
description: シングルサインオン (SSO) を使用するために、ASP.NET バックエンドで Office アドインを作成 (または変換) する方法に関するステップバイステップガイドです。
ms.date: 12/04/2019
localization_priority: Normal
ms.openlocfilehash: 71c5b6a90aa17ab08c1fe172be2181c9ec8650ef
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093722"
---
# <a name="create-an-aspnet-office-add-in-that-uses-single-sign-on-preview"></a><span data-ttu-id="e5c6e-103">シングル サインオンを使用する ASP.NET Office アドインを作成する (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="e5c6e-103">Create an ASP.NET Office Add-in that uses single sign-on (preview)</span></span>

<span data-ttu-id="e5c6e-104">ユーザーが Office にサインインしたとき、アドインは同じ資格情報を使用し、再度のサインインを要求することなく、複数のアプリケーションへのアクセスを許可することができます。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-104">When users are signed in to Office, your add-in can use the same credentials to permit users to access multiple applications without requiring them to sign in a second time.</span></span> <span data-ttu-id="e5c6e-105">概要については、「[Office アドインで SSO を有効化する](sso-in-office-add-ins.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-105">For an overview, see [Enable SSO in an Office Add-in](sso-in-office-add-ins.md).</span></span>
<span data-ttu-id="e5c6e-106">この記事では、Node.js と Express を使用して作成したアドインで、シングル サインオン (SSO) を有効化するプロセスについて手順を追って説明します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-106">This article walks you through the process of enabling single sign-on (SSO) in an add-in that is built with Node.js and Express.</span></span>

> [!NOTE]
> <span data-ttu-id="e5c6e-107">ASP.NET ベースのアドインに関する同様の記事については、「[シングル サインオンを使用する Node.js Office アドインを作成する](create-sso-office-add-ins-nodejs.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-107">For a similar article about an ASP.NET-based add-in, see [Create a Node.js Office Add-in that uses single sign-on](create-sso-office-add-ins-nodejs.md).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="e5c6e-108">前提条件</span><span class="sxs-lookup"><span data-stu-id="e5c6e-108">Prerequisites</span></span>

* <span data-ttu-id="e5c6e-109">Visual Studio 2019 以降。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-109">Visual Studio 2019 or later.</span></span>

* [<span data-ttu-id="e5c6e-110">Office Developer Tools</span><span class="sxs-lookup"><span data-stu-id="e5c6e-110">Office Developer Tools</span></span>](https://www.visualstudio.com/features/office-tools-vs.aspx)

[!include[additional prerequisites](../includes/sso-tutorial-prereqs.md)]

* <span data-ttu-id="e5c6e-111">少なくとも、Microsoft 365 サブスクリプションの OneDrive for Business に格納されているファイルとフォルダーがいくつかあります。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-111">At least a few files and folders stored on OneDrive for Business in your Microsoft 365 subscription.</span></span>

* <span data-ttu-id="e5c6e-112">Microsoft Azure サブスクリプション。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-112">A Microsoft Azure subscription.</span></span> <span data-ttu-id="e5c6e-113">このアドインには、Azure Active Directory (AD) が必要です。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-113">This add-in requires Azure Active Directory (AD).</span></span> <span data-ttu-id="e5c6e-114">Azure AD は、アプリケーションが認証および承認に使用する ID サービスを提供します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-114">Azure AD provides identity services that applications use for authentication and authorization.</span></span> <span data-ttu-id="e5c6e-115">[Microsoft Azure](https://account.windowsazure.com/SignUp) で試用版サブスクリプションを取得できます。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-115">A trial subscription can be acquired at [Microsoft Azure](https://account.windowsazure.com/SignUp).</span></span>

## <a name="set-up-the-starter-project"></a><span data-ttu-id="e5c6e-116">スタート プロジェクトをセットアップする</span><span class="sxs-lookup"><span data-stu-id="e5c6e-116">Set up the starter project</span></span>

<span data-ttu-id="e5c6e-117">「[Office Add-in ASPNET SSO](https://github.com/officedev/office-add-in-aspnet-sso)」にあるリポジトリを複製するかダウンロードします。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-117">Clone or download the repo at [Office Add-in ASPNET SSO](https://github.com/officedev/office-add-in-aspnet-sso).</span></span>

> [!NOTE]
> <span data-ttu-id="e5c6e-118">サンプルには 2 つのバージョンがあります。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-118">There are two versions of the sample:</span></span>
>
> * <span data-ttu-id="e5c6e-119">The **Before** folder is a starter project.</span><span class="sxs-lookup"><span data-stu-id="e5c6e-119">The **Before** folder is a starter project.</span></span> <span data-ttu-id="e5c6e-120">The UI and other aspects of the add-in that are not directly connected to SSO or authorization are already done.</span><span class="sxs-lookup"><span data-stu-id="e5c6e-120">The UI and other aspects of the add-in that are not directly connected to SSO or authorization are already done.</span></span> <span data-ttu-id="e5c6e-121">Later sections of this article walk you through the process of completing it.</span><span class="sxs-lookup"><span data-stu-id="e5c6e-121">Later sections of this article walk you through the process of completing it.</span></span>
> * <span data-ttu-id="e5c6e-122">このサンプルの **[Complete]** バージョンは、この記事の手順を完了したときに得られるアドインと同様のものですが、完成済みのプロジェクトには、この記事のテキストと重複するコード コメントが含まれています。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-122">The **Complete** version of the sample is just like the add-in that you would have if you completed the procedures of this article, except that the completed project has code comments that would be redundant with the text of this article.</span></span> <span data-ttu-id="e5c6e-123">完成済みのバージョンを使用する場合は、この記事の手順をそのまま実行しますが、[Before] を [Complete] に置き換えて、「**クライアント側のコードを作成する**」と「**サーバー側のコードを作成する**」のセクションを省略してください。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-123">To use the completed version, just follow the instructions in this article, but replace "Before" with "Complete" and skip the sections **Code the client side** and **Code the server side**.</span></span>


## <a name="register-the-add-in-with-azure-ad-v20-endpoint"></a><span data-ttu-id="e5c6e-124">Azure AD v2.0 エンドポイントにアドインを登録する</span><span class="sxs-lookup"><span data-stu-id="e5c6e-124">Register the add-in with Azure AD v2.0 endpoint</span></span>

1. <span data-ttu-id="e5c6e-125">[Azure ポータル - アプリの登録](https://go.microsoft.com/fwlink/?linkid=2083908)ページに移動してアプリを登録します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-125">Navigate to the [Azure portal - App registrations](https://go.microsoft.com/fwlink/?linkid=2083908) page to register your app.</span></span>

1. <span data-ttu-id="e5c6e-126">Microsoft 365 テナントに対して***管理者***の資格情報を使用してサインインします。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-126">Sign in with the ***admin*** credentials to your Microsoft 365 tenancy.</span></span> <span data-ttu-id="e5c6e-127">たとえば、MyName@contoso.onmicrosoft.com です。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-127">For example, MyName@contoso.onmicrosoft.com.</span></span>

1. <span data-ttu-id="e5c6e-128">**[新規登録]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-128">Select **New registration**.</span></span> <span data-ttu-id="e5c6e-129">**[アプリケーションを登録]** ページで、次のように値を設定します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-129">On the **Register an application** page, set the values as follows.</span></span>

    * <span data-ttu-id="e5c6e-130">`Office-Add-in-ASPNET-SSO` に **[名前]** を設定します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-130">Set **Name** to `Office-Add-in-ASPNET-SSO`.</span></span>
    * <span data-ttu-id="e5c6e-131">[**サポートされているアカウントの種類**] を [**任意の組織のディレクトリ内のアカウント (任意の Azure AD ディレクトリ - マルチテナント) と個人用の Microsoft アカウント (例: Skype、 Xbox)**] に設定します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-131">Set **Supported account types** to **Accounts in any organizational directory (Any Azure AD directory - Multitenant) and personal Microsoft accounts (e.g. Skype, Xbox)**.</span></span> <span data-ttu-id="e5c6e-132">(登録しているテナントのユーザーだけがアドインを使用できるようにする場合は、代わりに [**この組織ディレクトリのアカウントのみ...**] を選択します。ただし、追加セットアップ手順を実行する必要があります。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-132">(If you want the add-in to be usable only by users in the tenancy where you are registering it, you can choose **Accounts in this organizational directory only ...** instead, but you will need to go through some additional setup steps.</span></span> <span data-ttu-id="e5c6e-133">詳細については、「**シングルテナントのセットアップ**」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-133">See **Setup for single-tenant** below.)</span></span>
    * <span data-ttu-id="e5c6e-134">[**リダイレクト URI**] セクションで、ドロップダウンで [**Web**] が選択されていることを確認し、URI を [` https://localhost:44355/AzureADAuth/Authorize`] に設定します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-134">In the **Redirect URI** section, ensure that **Web** is selected in the drop down and then set the URI to` https://localhost:44355/AzureADAuth/Authorize`.</span></span>
    * <span data-ttu-id="e5c6e-135">**[登録]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-135">Choose **Register**.</span></span>

1. <span data-ttu-id="e5c6e-136">**Office-Add-in-NodeJS-SSO** ページで、**アプリケーション (クライアント) ID** と**ディレクトリ (テナント) ID** の値をコピーして保存します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-136">On the **Office-Add-in-NodeJS-SSO** page, copy and save the values for the **Application (client) ID** and the **Directory (tenant) ID**.</span></span> <span data-ttu-id="e5c6e-137">以降の手順では、それらの両方を使用します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-137">You'll use both of them in later procedures.</span></span>

    > [!NOTE]
    > <span data-ttu-id="e5c6e-138">この ID は、Office ホスト アプリケーション (たとえば、PowerPoint、Word、Excel) などの別のアプリケーションが、このアプリケーションへの承認されたアクセスを求めるときの「対象ユーザー」値になります。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-138">This ID is the "audience" value when other applications, such as the Office host application (e.g., PowerPoint, Word, Excel), seek authorized access to the application.</span></span> <span data-ttu-id="e5c6e-139">また、そのアプリケーションが Microsoft Graph への承認されたアクセスを求めるときには、このアプリケーションの「クライアント ID」になります。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-139">It is also the "client ID" of the application when it, in turn, seeks authorized access to Microsoft Graph.</span></span>

1. <span data-ttu-id="e5c6e-140">[**管理**] で [**証明書とシークレット**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-140">Under **Manage**, select **Certificates & secrets**.</span></span> <span data-ttu-id="e5c6e-141">[**新しいクライアント シークレット**] ボタンを選択します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-141">Select the **New client secret** button.</span></span> <span data-ttu-id="e5c6e-142">[**説明**] に値を入力してから、[**有効期限**] の適切なオプションを選択し、[**追加**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-142">Enter a value for **Description**, then select an appropriate option for **Expires** and choose **Add**.</span></span> <span data-ttu-id="e5c6e-143">後の手順で必要になるため、先に進む前に、*クライアント シークレットの値をすぐにコピーし、アプリケーション ID とともに保存*します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-143">*Copy the client secret value immediately and save it with the application ID* before proceeding as you'll need it in a later procedure.</span></span>

1. <span data-ttu-id="e5c6e-144">[**管理**] で [**API の公開**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-144">Under **Manage**, select **Expose an API**.</span></span> <span data-ttu-id="e5c6e-145">**[設定]** リンクを選択して、"api://$App ID GUID$" の形式でアプリケーション ID URI を生成します。$App ID GUID$ は**アプリケーション (クライアント) ID** です。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-145">Select the **Set** link to generate the Application ID URI in the form "api://$App ID GUID$", where $App ID GUID$ is the **Application (client) ID**.</span></span> <span data-ttu-id="e5c6e-146">`//` の後と GUID の前に、`localhost:44355/` を挿入します (末尾に追加されたスラッシュ「/」に注意します)。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-146">Insert `localhost:44355/` (note the forward slash "/" appended to the end) after the `//` and before the GUID.</span></span> <span data-ttu-id="e5c6e-147">ID 全体の形式は `api://localhost:44355/$App ID GUID$` でなければなりません (例: `api://localhost:44355/c6c1f32b-5e55-4997-881a-753cc1d563b7`)。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-147">The entire ID should have the form `api://localhost:44355/$App ID GUID$`; for example `api://localhost:44355/c6c1f32b-5e55-4997-881a-753cc1d563b7`.</span></span>

1. <span data-ttu-id="e5c6e-148">ダイアログで [**保存**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-148">Select **Save** on the dialog.</span></span>

1. <span data-ttu-id="e5c6e-149">**[Scope の追加]** ボタンをクリックします。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-149">Select the **Add a scope** button.</span></span> <span data-ttu-id="e5c6e-150">開いたパネルで、`access_as_user`を **[スコープ名]** として入力します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-150">In the panel that opens, enter `access_as_user` as the **Scope** name.</span></span>

1. <span data-ttu-id="e5c6e-151">**[同意できるのはだれですか?]** を **[管理者とユーザー]** に設定します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-151">Set **Who can consent?** to **Admins and users**.</span></span>

1. <span data-ttu-id="e5c6e-152">管理者およびユーザーの同意のプロンプトを構成するためのフィールドに、現在のユーザーと同じ権限で Office ホスト アプリケーションがアドインの Web API を使用できるようにする `access_as_user` 範囲に適した値を入力します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-152">Fill in the fields for configuring the admin and user consent prompts with values that are appropriate for the `access_as_user` scope which enables the Office host application to use your add-in's web APIs with the same rights as the current user.</span></span> <span data-ttu-id="e5c6e-153">提案:</span><span class="sxs-lookup"><span data-stu-id="e5c6e-153">Suggestions:</span></span>

    - <span data-ttu-id="e5c6e-154">**管理者の同意のタイトル**: Office はユーザーとして機能できます。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-154">**Admin consent title**: Office can act as the user.</span></span>
    - <span data-ttu-id="e5c6e-155">**管理者の同意の説明**: 現在のユーザーと同じ権限で Office がアドインの Web API を呼び出すことを可能にします。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-155">**Admin consent description**: Enable Office to call the add-in's web APIs with the same rights as the current user.</span></span>
    - <span data-ttu-id="e5c6e-156">**ユーザーの同意のタイトル**: Office は自分として機能できます。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-156">**User consent title**: Office can act as you.</span></span>
    - <span data-ttu-id="e5c6e-157">**管理者の同意の説明**: 自分と同じ権限で Office がアドインの Web API を呼び出すことを可能にします。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-157">**Admin consent description**: Enable Office to call the add-in's web APIs with the same rights that you have.</span></span>

1. <span data-ttu-id="e5c6e-158">**[状態]** が **[有効]** に設定されていることを確認してください。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-158">Ensure that **State** is set to **Enabled**.</span></span>

1. <span data-ttu-id="e5c6e-159">**[スコープの追加]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-159">Select **Add scope** .</span></span>

    > [!NOTE]
    > <span data-ttu-id="e5c6e-160">テキストフィールドのすぐ下に表示される **[スコープ名]** のドメイン部分は、以前に設定したアプリケーション ID URI に自動的に一致し、末尾に`/access_as_user`が追加されます。たとえば、`api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`です。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-160">The domain part of the **Scope** name displayed just below the text field should automatically match the Application ID URI that you set earlier, with `/access_as_user` appended to the end; for example, `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`.</span></span>

1. <span data-ttu-id="e5c6e-161">**[承認済みのクライアント アプリケーション]** セクションで、アドインの Web アプリケーションに対して承認するアプリケーションを特定します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-161">In the **Authorized client applications** section, you identify the applications that you want to authorize to your add-in's web application.</span></span> <span data-ttu-id="e5c6e-162">次のそれぞれの ID を事前承認する必要があります。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-162">Each of the following IDs needs to be pre-authorized.</span></span>

    - <span data-ttu-id="e5c6e-163">`d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)</span><span class="sxs-lookup"><span data-stu-id="e5c6e-163">`d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)</span></span>
    - <span data-ttu-id="e5c6e-164">`ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (Microsoft Office)</span><span class="sxs-lookup"><span data-stu-id="e5c6e-164">`ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (Microsoft Office)</span></span>
    - <span data-ttu-id="e5c6e-165">`57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office on the web)</span><span class="sxs-lookup"><span data-stu-id="e5c6e-165">`57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office on the web)</span></span>
    - <span data-ttu-id="e5c6e-166">`bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Outlook on the web)</span><span class="sxs-lookup"><span data-stu-id="e5c6e-166">`bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Outlook on the web)</span></span>

    <span data-ttu-id="e5c6e-167">ID ごとに、次の手順を実行します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-167">For each ID, take these steps:</span></span>

    <span data-ttu-id="e5c6e-168">a. </span><span class="sxs-lookup"><span data-stu-id="e5c6e-168">a.</span></span> <span data-ttu-id="e5c6e-169">**[クライアント アプリケーションの追加]** ボタンを選択し、表示されたパネルで [クライアント ID] をそれぞれの GUID に設定して、`api://localhost:44355/$App ID GUID$/access_as_user`のチェック ボックスをオンにします。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-169">Select **Add a client application** button and then, in the panel that opens, set the Client ID to the respective GUID and check the box for `api://localhost:44355/$App ID GUID$/access_as_user`.</span></span>

    <span data-ttu-id="e5c6e-170">b. </span><span class="sxs-lookup"><span data-stu-id="e5c6e-170">b.</span></span> <span data-ttu-id="e5c6e-171">**[アプリケーションの追加]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-171">Select **Add application**.</span></span>

1. <span data-ttu-id="e5c6e-172">[**管理**] で [**API のアクセス許可**]、[**アクセス許可の追加**] の順に選択します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-172">Under **Manage**, select **API permissions** and then select **Add a permission**.</span></span> <span data-ttu-id="e5c6e-173">開いたパネルで、**[Microsoft Graph]** を選択してから **[委任されたアクセス許可]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-173">On the panel that opens, choose **Microsoft Graph** and then choose **Delegated permissions**.</span></span>

1. <span data-ttu-id="e5c6e-174">アドインに必要な権限を検索するには、**[アクセス許可を選択]** の検索ボックスを使用します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-174">Use the **Select permissions** search box to search for the permissions your add-in needs.</span></span> <span data-ttu-id="e5c6e-175">以下を選択します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-175">Select the following.</span></span> <span data-ttu-id="e5c6e-176">アドイン自体に実際に必要なものは最初のもののみですが、Office ホストがアドインの Web アプリケーションへのトークンを取得するには、`profile` のアクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-176">Only the first is really required by your add-in itself; but the `profile` permission is required for the Office host to get a token to your add-in web application.</span></span> <span data-ttu-id="e5c6e-177">(実際には、Files.Read.All とプロファイルのみがアドインに必要です。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-177">(Only Files.Read.All and profile are actually needed by the add-in.</span></span> <span data-ttu-id="e5c6e-178">MSAL.NET ライブラリに必要なので、他の 2 つを要求する必要があります。)</span><span class="sxs-lookup"><span data-stu-id="e5c6e-178">You must request the other two because the MSAL.NET library requires them.)</span></span>

    * <span data-ttu-id="e5c6e-179">Files.Read.All</span><span class="sxs-lookup"><span data-stu-id="e5c6e-179">Files.Read.All</span></span>
    * <span data-ttu-id="e5c6e-180">offline_access</span><span class="sxs-lookup"><span data-stu-id="e5c6e-180">offline_access</span></span>
    * <span data-ttu-id="e5c6e-181">openid</span><span class="sxs-lookup"><span data-stu-id="e5c6e-181">openid</span></span>
    * <span data-ttu-id="e5c6e-182">profile</span><span class="sxs-lookup"><span data-stu-id="e5c6e-182">profile</span></span>

    > [!NOTE]
    > <span data-ttu-id="e5c6e-183">`User.Read` アクセス許可は既定でリストされています。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-183">The `User.Read` permission may already be listed by default.</span></span> <span data-ttu-id="e5c6e-184">必要でないアクセス許可は依頼しない方がよいため、アドインが実際に必要でない場合は、このアクセス許可のボックスのチェックをオフにしておくことをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-184">It is a good practice not to ask for permissions that are not needed, so we recommend that you uncheck the box for this permission if your add-in does not actually need it.</span></span>

1. <span data-ttu-id="e5c6e-185">表示される各アクセス許可のチェック ボックスをオンにします。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-185">Select the check box for each permission as it appears.</span></span> <span data-ttu-id="e5c6e-186">アドインに必要なアクセス許可を選択したら、パネルの下部にある **[アクセス許可を追加する]** ボタンをクリックします。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-186">After selecting the permissions that your add-in needs, select the **Add permissions** button at the bottom of the panel.</span></span>

1. <span data-ttu-id="e5c6e-187">同じページで、[**[テナント名] に管理者の同意を与える**] ボタンを選択し、表示される確認に対して [**同意する**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-187">On the same page, choose the **Grant admin consent for [tenant name]** button, and then select **Accept** for the confirmation that appears.</span></span>

    > [!NOTE]
    > <span data-ttu-id="e5c6e-188">[**[テナント名] に管理者の同意を与える**] を選択すると、同意プロンプトを作成できるように、数分後に再試行を求めるバナー メッセージが表示される場合があります。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-188">After choosing **Grant admin consent for [tenant name]**, you may see a banner message asking you to try again in a few minutes so that the consent prompt can be constructed.</span></span> <span data-ttu-id="e5c6e-189">その場合は、次のセクションで作業を開始できますが、***必ずポータルに戻り、このボタンを押してください***。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-189">If so, you can start work on the next section, ***but don't forget to come back to the portal and press this button***!</span></span>

## <a name="configure-the-solution"></a><span data-ttu-id="e5c6e-190">ソリューションを構成する</span><span class="sxs-lookup"><span data-stu-id="e5c6e-190">Configure the solution</span></span>

1. <span data-ttu-id="e5c6e-191">[**Before**] フォルダーのルートで、**Visual Studio** でソリューション (.sln) ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-191">In the root of the **Before** folder, open the solution (.sln) file in **Visual Studio**.</span></span> <span data-ttu-id="e5c6e-192">[**ソリューション エクスプローラー**] の一番上のノード (プロジェクト ノードではなく、ソリューション ノード) を右クリックして、[**スタートアップ プロジェクトの設定**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-192">Right-click the top node in **Solution Explorer** (the Solution node, not either of the project nodes), and then select **Set startup projects**.</span></span>

1. <span data-ttu-id="e5c6e-193">[**共通プロパティ**] で、[**スタートアップ プロジェクト**]、[**マルチ スタートアップ プロジェクト**] の順に選択します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-193">Under **Common Properties**, select **Startup Project**, and then **Multiple startup projects**.</span></span> <span data-ttu-id="e5c6e-194">両方のプロジェクトの [**アクション**] が [**開始**] に設定され、「... WebAPI」で終わるプロジェクトが最初にリストされていることを確認します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-194">Ensure that the **Action** for both projects is set to **Start**, and that the project that ends in "...WebAPI" is listed first.</span></span> <span data-ttu-id="e5c6e-195">ダイアログを閉じます。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-195">Close the dialog.</span></span>

1. <span data-ttu-id="e5c6e-196">[**ソリューション エクスプローラー**] に戻り、[**Office-Add-in-Microsoft-Graph-ASPNETWebAPI**] プロジェクトを選択します (右クリックしないでください)。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-196">Back in **Solution Explorer**, select (don't right-click) the **Office-Add-in-Microsoft-Graph-ASPNETWebAPI** project.</span></span> <span data-ttu-id="e5c6e-197">[**プロパティ**] ウィンドウを開きます。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-197">The **Properties** pane opens.</span></span> <span data-ttu-id="e5c6e-198">[**SSL 有効**] が [**True**] であることを確認します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-198">Ensure that **SSL Enabled** is **True**.</span></span> <span data-ttu-id="e5c6e-199">[**SSL URL**] が `http://localhost:44355/` であることを確認します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-199">Verify that the **SSL URL** is `http://localhost:44355/`.</span></span>

1. <span data-ttu-id="e5c6e-200">「Web.config」 で、以前にコピーした値を使用します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-200">In "Web.config", use the values that you copied in earlier.</span></span> <span data-ttu-id="e5c6e-201">[**ida:ClientID**] と [**ida:Audience**] の両方を [**アプリケーション (クライアント) ID**] に設定し、[**ida:Password**] をクライアント シークレットに設定します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-201">Set both the **ida:ClientID** and the **ida:Audience** to your **Application (client) ID**, and set **ida:Password** to your client secret.</span></span>

    > [!NOTE]
    > <span data-ttu-id="e5c6e-202">[**アプリケーション (クライアント) ID**] は、Office ホスト アプリケーション (たとえば、PowerPoint、Word、Excel) などの別のアプリケーションが、このアプリケーションへの承認されたアクセスを求めるときの「対象ユーザー」値になります。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-202">The **Application (client) ID** is the "audience" value when other applications, such as the Office host application (e.g., PowerPoint, Word, Excel), seek authorized access to the application.</span></span> <span data-ttu-id="e5c6e-203">また、そのアプリケーションが Microsoft Graph への承認されたアクセスを求めるときには、このアプリケーションの「クライアント ID」になります。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-203">It is also the "client ID" of the application when it, in turn, seeks authorized access to Microsoft Graph.</span></span>

1. <span data-ttu-id="e5c6e-204">アドインを登録したときに、**サポートされているアカウントの種類**で「この組織のディレクトリ内のアカウントのみ」を選択しなかった場合は、web.config を保存して閉じます。 それ以外の場合は、保存して、開いたままにします。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-204">If you didn't choose "Accounts in this organizational directory only" for **SUPPORTED ACCOUNT TYPES** when you registered the add-in, save and close the web.config. Otherwise, save but leave it open.</span></span>

1. <span data-ttu-id="e5c6e-205">[**ソリューション エクスプローラー**] で [**Office-Add-in-Microsoft-Graph-ASPNET**] プロジェクトを選択し、アドイン マニフェスト ファイル「Office-Add-in-ASPNET-SSO.xml」を開いて、ファイルの下部までスクロールします。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-205">Still in **Solution Explorer**, choose the **Office-Add-in-Microsoft-Graph-ASPNET** project and open the add-in manifest file “Office-Add-in-ASPNET-SSO.xml” and then scroll to the bottom of the file.</span></span> <span data-ttu-id="e5c6e-206">`</VersionOverrides>` 終了タグの直前に、以下のマークアップがあります。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-206">Just above the end `</VersionOverrides>` tag, you'll find the following markup:</span></span>

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

1. <span data-ttu-id="e5c6e-207">このマークアップ内の*両方の場所の*プレースホルダー「$application_GUID here$」を、アドインの登録時にコピーしたアプリケーション ID に置き換えます。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-207">Replace the placeholder “$application_GUID here$” *in both places* in the markup with the Application ID that you copied when you registered your add-in.</span></span> <span data-ttu-id="e5c6e-208">「$」は ID の一部ではないので、これらを含めないでください。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-208">The "$" signs are not part of the ID, so do not include them.</span></span> <span data-ttu-id="e5c6e-209">これは、web.config の ClientID と Audience に使用したものと同じ ID です。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-209">This is the same ID you used in for the ClientID and Audience in the web.config.</span></span>

  > [!NOTE]
  > <span data-ttu-id="e5c6e-210">**リソース**値は、アドインを登録したときに設定した**アプリケーション ID URI** です。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-210">The **Resource** value is the **Application ID URI** you set when you registered the add-in.</span></span> <span data-ttu-id="e5c6e-211">**[範囲]** セクションは、アドインが AppSource を通じて販売される場合に同意ダイアログ ボックスを生成するためにのみ使用されます。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-211">The **Scopes** section is used only to generate a consent dialog box if the add-in is sold through AppSource.</span></span>

1. <span data-ttu-id="e5c6e-212">ファイルを保存して閉じます。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-212">Save and close the file.</span></span>

### <a name="setup-for-single-tenant"></a><span data-ttu-id="e5c6e-213">シングルテナントのセットアップ</span><span class="sxs-lookup"><span data-stu-id="e5c6e-213">Setup for single-tenant</span></span>

<span data-ttu-id="e5c6e-214">アドインを登録したときに、**サポートされているアカウントの種類**で「この組織のディレクトリ内のアカウントのみ」を選択した場合は、これらの追加のセットアップ手順を実行する必要があります。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-214">If you chose "Accounts in this organizational directory only" for **SUPPORTED ACCOUNT TYPES** when you registered the add-in, you need to take these additional setup steps:</span></span>

1. <span data-ttu-id="e5c6e-215">Azure ポータルに戻り、アドインの登録の [**概要**] ブレードを開きます。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-215">Go back to the Azure Portal and open the **Overview** blade of the add-in's registration.</span></span> <span data-ttu-id="e5c6e-216">[**Directory (テナント) ID**] をコピーします。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-216">Copy the **Directory (tenant) ID**.</span></span>

1. <span data-ttu-id="e5c6e-217">web.config で、[**ida：Authority**] の値の「Common」を前の手順でコピーした GUID に置き換えます。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-217">In the web.config, replace the "common" in the value of **ida:Authority** with the GUID you copied in the preceding step.</span></span> <span data-ttu-id="e5c6e-218">終了すると、値は `<add key="ida:Authority" value="https://login.microsoftonline.com/12345678-91ab-cdef-0123-456789abcdef/oauth2/v2.0" />` のようになります。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-218">When you are finished the value should look similar to this: `<add key="ida:Authority" value="https://login.microsoftonline.com/12345678-91ab-cdef-0123-456789abcdef/oauth2/v2.0" />`.</span></span>

1. <span data-ttu-id="e5c6e-219">web.config を保存して閉じます。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-219">Save and close the web.config.</span></span>

## <a name="code-the-client-side"></a><span data-ttu-id="e5c6e-220">クライアント側のコードの作成</span><span class="sxs-lookup"><span data-stu-id="e5c6e-220">Code the client side</span></span>

1. <span data-ttu-id="e5c6e-221">[**スクリプト**] フォルダー内の HomeES6.js ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-221">Open the HomeES6.js file in the **Scripts** folder.</span></span> <span data-ttu-id="e5c6e-222">これには、一部のコードが既に含まれています。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-222">It already has some code in it:</span></span>

    * <span data-ttu-id="e5c6e-223">Office が UI に Internet Explorer を使用しているときにアドインを実行できるように、Office.Promise オブジェクトをグローバル ウィンドウ オブジェクトに割り当てるポリフィル。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-223">A polyfill that assigns the Office.Promise object to the global window object so that the add-in can run when Office is using Internet Explorer for the UI.</span></span> <span data-ttu-id="e5c6e-224">(詳細については、「[Office アドインによって使用されるブラウザー](../concepts/browsers-used-by-office-web-add-ins.md)」を参照してください。)</span><span class="sxs-lookup"><span data-stu-id="e5c6e-224">(For more information, see [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).)</span></span>
    * <span data-ttu-id="e5c6e-225">`Office.initialize` メソッドへの割り当てが、`getGraphAccessTokenButton` ボタン クリック イベントへのハンドラーの割り当てになります。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-225">An assignment to the `Office.initialize` method that, in turn, assigns a handler to the `getGraphAccessTokenButton` button click event.</span></span>
    * <span data-ttu-id="e5c6e-226">`showResult` メソッドは、作業ウィンドウの下側に Microsoft Graph から返されたデータ (またはエラー メッセージ) を表示するものです。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-226">A `showResult` method that will display data returned from Microsoft Graph (or an error message) at the bottom of the task pane.</span></span>
    * <span data-ttu-id="e5c6e-227">`logErrors` メソッドは、エンド ユーザーを対象としていないエラーをコンソールにログ出力するものです。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-227">A `logErrors` method that will log to console errors that are not intended for the end user.</span></span>
    * <span data-ttu-id="e5c6e-228">SSO がサポートされていない、または SSO がエラーになっているシナリオでアドインが使用するフォールバック認証システムを実装するコード。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-228">Code that implements the fallback authorization system that the add-in will use in scenarios where SSO is not supported or has errored.</span></span>

1. <span data-ttu-id="e5c6e-229">`Office.initialize` への割り当ての下に、次に示すコードを追加します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-229">Below the assignment to `Office.initialize`, add the code below.</span></span> <span data-ttu-id="e5c6e-230">このコードについては、次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-230">Note the following about this code:</span></span>

    * <span data-ttu-id="e5c6e-231">アドインのエラー処理により、アクセス トークンの取得が別のオプションのセットを使用して自動的に再試行されることがあります。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-231">The error-handling in the add-in will sometimes automatically attempt a second time to get an access token, using a different set of options.</span></span> <span data-ttu-id="e5c6e-232">カウンター変数 `retryGetAccessToken` は、ユーザーがトークンを取得しようとしたときに繰り返し再試行されないように使用されます。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-232">The counter variable `retryGetAccessToken` is used to ensure that the user isn't cycled repeatedly through failed attempts to get a token.</span></span>
    * <span data-ttu-id="e5c6e-233">`getGraphData` 関数は、ES6 `async` キーワードで定義されます。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-233">The `getGraphData` function is defined with the ES6 `async` keyword.</span></span> <span data-ttu-id="e5c6e-234">ES6 構文を使用すると、Office アドインの SSO API の使用が非常に簡単になります。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-234">Using ES6 syntax makes the SSO API in Office Add-ins much easier to to use.</span></span> <span data-ttu-id="e5c6e-235">これは、ソリューション内の、Internet Explorer でサポートされていない構文を使用する唯一のファイルです。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-235">This is the only file in the solution that will use syntax that is not supported by Internet Explorer.</span></span> <span data-ttu-id="e5c6e-236">ファイル名に「ES6」というリマインダーが設定されています。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-236">We put 'ES6' in the filename as a reminder.</span></span> <span data-ttu-id="e5c6e-237">このソリューションでは、tsc トランスパイラーを使用してこのファイルを ES5 にトランスパイルします。これにより、Office が UI に Internet Explorer を使用しているときにアドインが実行されます。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-237">The solution uses the tsc transpiler to transpile this file to ES5, so that the add-in can run when Office is using Internet Explorer for the UI.</span></span> <span data-ttu-id="e5c6e-238">(プロジェクトのルートにある tsconfig.json ファイルを参照します。)</span><span class="sxs-lookup"><span data-stu-id="e5c6e-238">(See the tsconfig.json file in the root of the project.)</span></span>

    ```javascript
    var retryGetAccessToken = 0;

    async function getGraphData() {
        await getDataWithToken({ allowSignInPrompt: true, forMSGraphAccess: true });
    }
    ```

1. <span data-ttu-id="e5c6e-239">`getGraphData` 関数の下に、次の関数を追加します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-239">Below the `getGraphData` function add the following function.</span></span> <span data-ttu-id="e5c6e-240">後の手順で `handleClientSideErrors` 関数を作成することに注意してください。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-240">Note that you create the `handleClientSideErrors` function in a later step.</span></span>

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

1. <span data-ttu-id="e5c6e-241">`TODO 1`を以下のように置き換えます。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-241">Replace `TODO 1` with the following.</span></span> <span data-ttu-id="e5c6e-242">このコードについては、以下の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-242">About this code, note:</span></span>

    * <span data-ttu-id="e5c6e-243">`getAccessToken` は、Azure AD からブートストラップ トークンを取得し、アドインに戻るように Office に指示します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-243">`getAccessToken` tells Office to get a bootstrap token from Azure AD and return to the add-in.</span></span>
    * <span data-ttu-id="e5c6e-244">`allowSignInPrompt` は、ユーザーがまだ Office にサインインしていない場合、ユーザーにサインインするように求めるように Office に指示します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-244">`allowSignInPrompt` tells Office to prompt the user to sign in if the user isn't already signed into Office.</span></span>
    * <span data-ttu-id="e5c6e-245">`forMSGraphAccess` は、アドインが (ブートストラップ トークンをユーザー ID トークンとして使用するだけでなく) Microsoft Graph へのアクセス トークンのブートストラップ トークンを交換することを Office に通知します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-245">`forMSGraphAccess` tells Office that the add-in intends to swap the bootstrap token for an access token to Microsoft Graph (instead of just using the bootstrap token as a user ID token).</span></span> <span data-ttu-id="e5c6e-246">このオプションを設定すると、ユーザーのテナント管理者がアドインの同意を与えていない場合、Office はブートストラップ トークンの取得プロセスをキャンセルすることができます (そしてエラー コード 13012 が返されます)。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-246">Setting this option gives Office a chance to cancel the process of getting a bootstrap token (and return error code 13012) if the user's tenant administrator has not granted consent to the add-in.</span></span> <span data-ttu-id="e5c6e-247">アドインのクライアント側コードが 13012 に返信するには、フォールバック認証システムに分岐します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-247">The add-in's client-side code can respond to the 13012 by branching to a fallback authorization system.</span></span> <span data-ttu-id="e5c6e-248">`forMSGraphAccess` が使用されず、管理者が同意を与えていない場合は、ブートストラップ トークンが返されますが、on-behalf-of フローと交換しようとするとエラーになります。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-248">If the `forMSGraphAccess` is not used, and the admin has not granted consent, the bootstrap token is returned, but the attempt to exhange it with the on-behalf-of flow would result in an error.</span></span> <span data-ttu-id="e5c6e-249">したがって、`forMSGraphAccess` オプションを使用すると、アドインがフォールバック システムにすばやく分岐できます。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-249">Thus, the `forMSGraphAccess` option enables the add-in to branch to the fallback system quickly.</span></span>
    * <span data-ttu-id="e5c6e-250">後の手順で `getData` 関数を作成します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-250">You create the `getData` function in a later step.</span></span>
    * <span data-ttu-id="e5c6e-251">`/api/values` パラメーターは、トークンを交換したり、Microsoft Graph を呼び出すためにアクセス トークンを使用したりする、サーバー側コントローラーの URL です。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-251">The `/api/values` parameter is the URL of a server-side controller that will make the token exchange and use the access token it gets back to make the call to Microsoft Graph.</span></span>

    ```javascript
    let bootstrapToken = await OfficeRuntime.auth.getAccessToken({
        allowSignInPrompt: true,
        forMSGraphAccess: true });

    getData("/api/values", bootstrapToken);
    ```

1. <span data-ttu-id="e5c6e-252">`getGraphData` 関数の下に、次を追加します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-252">Below the `getGraphData` function, add the following.</span></span> <span data-ttu-id="e5c6e-253">このコードについては、以下の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-253">About this code, note:</span></span>

    * <span data-ttu-id="e5c6e-254">これは、SSO 認証システムおよびフォールバック認証システムの両方で使用されます。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-254">It is used by both the SSO and the fallback authorization systems.</span></span>
    * <span data-ttu-id="e5c6e-255">`relativeUrl` パラメーターは、サーバー側のコントローラーです。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-255">The `relativeUrl` parameter is a server-side controller.</span></span>
    * <span data-ttu-id="e5c6e-256">`accessToken` パラメーターは、ブートストラップ トークンまたはフル アクセス トークンにすることができます。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-256">The `accessToken` parameter can be a bootstrap token or a full access token.</span></span>
    * <span data-ttu-id="e5c6e-257">`writeFileNamesToOfficeDocument` は、既にプロジェクトの一部です。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-257">The `writeFileNamesToOfficeDocument` is already part of the project.</span></span>
    * <span data-ttu-id="e5c6e-258">後の手順で `handleServerSideErrors` 関数を作成します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-258">You create the `handleServerSideErrors` function in a later step.</span></span>

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

### <a name="handle-client-side-errors"></a><span data-ttu-id="e5c6e-259">クライアント側のエラーを処理する</span><span class="sxs-lookup"><span data-stu-id="e5c6e-259">Handle client-side errors</span></span>

1. <span data-ttu-id="e5c6e-260">`getData`関数の下に、次の関数を追加します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-260">Below the `getData` function, add the following function.</span></span> <span data-ttu-id="e5c6e-261">`error.code`は数値であり、通常は 13xxx の範囲にあることを注意してください。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-261">Note that `error.code` is a number, usually in the range 13xxx.</span></span>

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

1. <span data-ttu-id="e5c6e-262">`TODO 2`を以下のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-262">Replace `TODO 2` with the following code.</span></span> <span data-ttu-id="e5c6e-263">これらのエラーの詳細については、「[Office アドインの SSO のトラブルシューティング (Troubleshoot SSO in Office Add-ins)](troubleshoot-sso-in-office-add-ins.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-263">For more information about these errors, see [Troubleshoot SSO in Office Add-ins](troubleshoot-sso-in-office-add-ins.md).</span></span>

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
        // Only seen in Office on the web.
        showResult(["Office on the web is experiencing a problem. Please sign out of Office, close the browser, and then start again."]);
        break;
    case 13008:
        // Only seen in Office on the web.
        showResult(["Office is still working on the last operation. When it completes, try this operation again."]);
        break;
    case 13010:
        // Only seen in Office on the web.
        showResult(["Follow the instructions to change your browser's zone configuration."]);
        break;
    ```

1. <span data-ttu-id="e5c6e-264">`TODO 3`を以下のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-264">Replace `TODO 3` with the following code.</span></span> <span data-ttu-id="e5c6e-265">その他のエラーが発生した場合、アドインはフォールバック認証システムに分岐します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-265">For all other errors, the add-in branches to the fallback authorization system.</span></span> <span data-ttu-id="e5c6e-266">これらのエラーの詳細については、「[Office アドインの SSO のトラブルシューティング](troubleshoot-sso-in-office-add-ins.md)」を参照してください。このアドインでは、フォールバック システムはユーザーが既にログインしている場合でもユーザーのサインインを要求するダイアログを開き、msal.js および暗黙的フローを使用して Microsoft Graph へのアクセス トークンを取得します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-266">For more information about these errors, see [Troubleshoot SSO in Office Add-ins](troubleshoot-sso-in-office-add-ins.md). In this add-in, the fallback system opens a dialog which requires the user to sign in, even if the user already is, and uses msal.js and the Implicit Flow to get an access token to Microsoft Graph.</span></span>

    ```javascript
    default:
        dialogFallback();
        break;
    ```

### <a name="handle-server-side-errors"></a><span data-ttu-id="e5c6e-267">サーバー側のエラーを処理する</span><span class="sxs-lookup"><span data-stu-id="e5c6e-267">Handle server-side errors</span></span>

1. <span data-ttu-id="e5c6e-268">`handleClientSideErrors` 関数の下に、次の関数を追加します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-268">Below the `handleClientSideErrors` function, add the following function.</span></span>

    ```javascript
    function handleServerSideErrors(result) {

    // TODO 4: Parse the JSON response.

    // TODO 5: Handle case where Microsoft Graph requires an additional form
    //         of authentication.

    // TODO 6: Handle other Azure AD errors

    }
    ```

1. <span data-ttu-id="e5c6e-269">`TODO 4`を以下のように置き換えます。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-269">Replace `TODO 4` with the following.</span></span> <span data-ttu-id="e5c6e-270">このコードについては、MFA などが存在する前に ASP.NET エラー クラスが作成されたことに注意してください。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-270">About this code, note that ASP.NET error classes were created before there was such a thing as MFA.</span></span> <span data-ttu-id="e5c6e-271">第 2 認証要素に対する要求をサーバー側の論理が処理する方法の副作用として、クライアントに送信されるサーバー側のエラーは **Message** プロパティがありますが、**ExceptionMessage** プロパティはありません。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-271">As a side-effect of how our server-side logic handles the requests for a second authentication factor, the server-side error sent to the client has a **Message** property but no **ExceptionMessage** property.</span></span> <span data-ttu-id="e5c6e-272">ただし、他のすべてのエラーには **ExceptionMessage** プロパティがあるため、クライアント側のコードは両方の応答を解析する必要があります。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-272">But all other errors will have a **ExceptionMessage** property, so the client-side code has to parse the response for both.</span></span> <span data-ttu-id="e5c6e-273">どちらか一方の変数が未定義になります。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-273">Either one or the other variable will be undefined.</span></span>

    ```javascript
    var message = JSON.parse(result.responseText).Message;
    var exceptionMessage = JSON.parse(result.responseText).ExceptionMessage;
    ```

1. <span data-ttu-id="e5c6e-274">`TODO 5`を以下のように置き換えます。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-274">Replace `TODO 5` with the following.</span></span> <span data-ttu-id="e5c6e-275">Microsoft Graph が認証の追加形式を必要とする場合、エラー AADSTS50076 が送信されます。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-275">When Microsoft Graph requires an additional form of authentication, it sends error AADSTS50076.</span></span> <span data-ttu-id="e5c6e-276">これには、**Message.Claims** プロパティの追加要件に関する情報が含まれます。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-276">It includes information about the additional requirement in the **Message.Claims** property.</span></span> <span data-ttu-id="e5c6e-277">これを処理するために、コードはブートストラップ トークンの取得を 2 回試行しますが、今回は `authChallenge` オプションの値として追加要素の要求が含まれます。これにより、Azure AD は、必要なすべての形式の認証をユーザーに要求します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-277">To handle this, the code makes a second attempt to get the bootstrap token, but this time it includes the request for an additional factor as the value of the `authChallenge` option, which tells Azure AD to prompt the user for all required forms of authentication.</span></span>

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

1. <span data-ttu-id="e5c6e-278">`TODO 6`を以下のように置き換えます。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-278">Replace `TODO 6` with the following.</span></span>

    ```javascript
    if (exceptionMessage) {

        // TODO 7: Handle case where bootstrap token has expired.

        // TODO 8: Handle all other Azure AD errors.
    }
    ```

1. <span data-ttu-id="e5c6e-279">`TODO 7`を以下のように置き換えます。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-279">Replace `TODO 7` with the following.</span></span> <span data-ttu-id="e5c6e-280">まれにブートストラップ トークンが Office の検証時に期限切れにならず、交換のために Azure AD に送信されるまでの間に期限切れになることがあることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-280">Note that on rare occasions the bootstrap token is unexpired when Office validates it, but expires by the time it is sent to Azure AD for exchange.</span></span> <span data-ttu-id="e5c6e-281">Azure AD はエラー AADSTS500133 で応答します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-281">Azure AD will respond with error AADSTS500133.</span></span> <span data-ttu-id="e5c6e-282">この場合、コードは SSO API を呼び戻します (ただし、1 回のみ)。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-282">When this happens, the code  recalls the SSO API (but no more than once).</span></span> <span data-ttu-id="e5c6e-283">今回は、Office が期限切れになっていない新しいブートストラップ トークンを返します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-283">This time Office returns a new unexpired bootstrap token.</span></span>

    ```javascript
    if ((exceptionMessage.indexOf("AADSTS500133") !== -1)
        && (retryGetAccessToken <= 0)) {

        retryGetAccessToken++;
        getGraphData();
    }
    ```

1. <span data-ttu-id="e5c6e-284">`TODO 8` を以下のように置き換えます。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-284">Replace `TODO 8` with the following.</span></span>

    ```javascript
    else {
        dialogFallback();
    }
    ```

1. <span data-ttu-id="e5c6e-285">ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-285">Save the file.</span></span>

## <a name="code-the-server-side"></a><span data-ttu-id="e5c6e-286">サーバー側のコードを作成する</span><span class="sxs-lookup"><span data-stu-id="e5c6e-286">Code the server side</span></span>

### <a name="configure-the-owin-middleware"></a><span data-ttu-id="e5c6e-287">OWIN ミドルウェアを構成する</span><span class="sxs-lookup"><span data-stu-id="e5c6e-287">Configure the OWIN middleware</span></span>

1. <span data-ttu-id="e5c6e-288">**Office-Add-in-ASPNET-SSO-WebAPI** プロジェクトのルートにある Startup.cs ファイルを開き、**スタートアップ** クラスに次のメソッドを追加します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-288">Open the Startup.cs file in the root of the **Office-Add-in-ASPNET-SSO-WebAPI** project and add the following method to the **Startup** class.</span></span> <span data-ttu-id="e5c6e-289">`ConfigureAuth` メソッドは、この後の手順で作成することに注意してください。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-289">Note that you create the `ConfigureAuth` method in a later step.</span></span>

    ```csharp
    public void Configuration(IAppBuilder app)
    {
        ConfigureAuth(app);
    }
    ```

1. <span data-ttu-id="e5c6e-290">ファイルを保存して閉じます。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-290">Save and close the file.</span></span>

1. <span data-ttu-id="e5c6e-291">**App_Start** フォルダーを右クリックして、**[追加] > [クラス]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-291">Right-click the **App_Start** folder and select **Add > Class**.</span></span>

1. <span data-ttu-id="e5c6e-292">**[新しい項目の追加]** ダイアログで、ファイルに「**Startup.Auth.cs**」という名前を付けて **[追加]** をクリックします。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-292">In the **Add new item** dialog name the file **Startup.Auth.cs** and then click **Add**.</span></span>

1. <span data-ttu-id="e5c6e-293">新しいファイルで名前空間の名前を `Office_Add_in_ASPNET_SSO_WebAPI` に短縮します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-293">Shorten the namespace name in the new file to `Office_Add_in_ASPNET_SSO_WebAPI`.</span></span>

1. <span data-ttu-id="e5c6e-294">ファイルの先頭に、次に示す `using` ステートメントがすべて揃っていることを確認します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-294">Ensure that all of the following `using` statements are at the top of the file.</span></span>

    ```csharp
    using Owin;
    using Microsoft.IdentityModel.Tokens;
    using System.Configuration;
    using Microsoft.Owin.Security.OAuth;
    using Microsoft.Owin.Security.Jwt;
    using Office_Add_in_ASPNET_SSO_WebAPI.App_Start;
    ```

1. <span data-ttu-id="e5c6e-295">Add the keyword `partial` to the declaration of the `Startup` class, if it is not already there.</span><span class="sxs-lookup"><span data-stu-id="e5c6e-295">Add the keyword `partial` to the declaration of the `Startup` class, if it is not already there.</span></span> <span data-ttu-id="e5c6e-296">It should look like this:</span><span class="sxs-lookup"><span data-stu-id="e5c6e-296">It should look like this:</span></span>

    `public partial class Startup`

1. <span data-ttu-id="e5c6e-297">Add the following method to the `Startup` class.</span><span class="sxs-lookup"><span data-stu-id="e5c6e-297">Add the following method to the `Startup` class.</span></span> <span data-ttu-id="e5c6e-298">This method specifies how the OWIN middleware will validate the access tokens that are passed to it from the `getData` method in the client-side Home.js file.</span><span class="sxs-lookup"><span data-stu-id="e5c6e-298">This method specifies how the OWIN middleware will validate the access tokens that are passed to it from the `getData` method in the client-side Home.js file.</span></span> <span data-ttu-id="e5c6e-299">The authorization process is triggered whenever a Web API endpoint that is decorated with the `[Authorize]` attribute is called.</span><span class="sxs-lookup"><span data-stu-id="e5c6e-299">The authorization process is triggered whenever a Web API endpoint that is decorated with the `[Authorize]` attribute is called.</span></span>

    ```csharp
    public void ConfigureAuth(IAppBuilder app)
    {
        // TODO 1: Configure the validation settings

        // TODO 2: Specify the type of authorization and the discovery endpoint
        //        of the secure token service.
    }
    ```

1. <span data-ttu-id="e5c6e-300">`TODO 1` を以下のように置き換えます。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-300">Replace the `TODO 1` with the following.</span></span> <span data-ttu-id="e5c6e-301">このコードの注意点は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-301">Note about this code:</span></span>

    * <span data-ttu-id="e5c6e-302">このコードでは、Office ホストから得られるブートストラップ トークンで指定された対象ユーザーが web.config で指定された値と一致する必要があることを OWIN に指示します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-302">The code instructs OWIN to ensure that the audience specified in the bootstrap token that comes from the Office host must match the value specified in the web.config.</span></span>
    * <span data-ttu-id="e5c6e-303">Microsoft アカウントには、組織のテナント GUID とは異なる発行者 GUID があります。そのため、両方の種類のアカウントをサポートするために、発行者は検証されません。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-303">Microsoft Accounts have an issuer GUID that is different from any organizational tenant GUID, so to support both kinds of accounts, we do not validate the issuer.</span></span>
    * <span data-ttu-id="e5c6e-304">`SaveSigninToken` を `true` に設定することで、OWIN は Office ホストからの生のブートストラップ トークンを保存するようになります。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-304">Setting `SaveSigninToken` to `true` causes OWIN to save the raw bootstrap token from the Office host.</span></span> <span data-ttu-id="e5c6e-305">これは、アドインが代理フローで Microsoft Graph へのアクセス トークンを取得するために必要になります。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-305">The add-in needs it to obtain an access token to Microsoft Graph with the on-behalf-of flow.</span></span>
    * <span data-ttu-id="e5c6e-306">OWIN ミドルウェアでは、スコープは検証されません。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-306">Scopes are not validated by the OWIN middleware.</span></span> <span data-ttu-id="e5c6e-307">`access_as_user` が含まれている必要があるブートストラップ トークンのスコープは、コントローラーで検証されます。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-307">The scopes of the bootstrap token, which should include `access_as_user`, is validated in the controller.</span></span>

    ```csharp
    TokenValidationParameters tvps = new TokenValidationParameters
    {
        ValidAudience = ConfigurationManager.AppSettings["ida:Audience"],
        ValidateIssuer = false,
        SaveSigninToken = true
    };
    ```

1. <span data-ttu-id="e5c6e-308">`TODO 2`を以下のように置き換えます。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-308">Replace `TODO 2` with the following.</span></span> <span data-ttu-id="e5c6e-309">このコードの注意点は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-309">Note about this code:</span></span>

    * <span data-ttu-id="e5c6e-310">より一般的な `UseWindowsAzureActiveDirectoryBearerAuthentication` は Azure AD V2 エンドポイントに準拠していないため、その代わりとしてメソッド `UseOAuthBearerAuthentication` が呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-310">The method `UseOAuthBearerAuthentication` is called instead of the more common `UseWindowsAzureActiveDirectoryBearerAuthentication` because the latter is not compatible with the Azure AD V2 endpoint.</span></span>
    * <span data-ttu-id="e5c6e-311">このメソッドに渡される URL は、Office ホストから受け取ったブートストラップ トークンの署名の検証に必要になるキーを取得するための方法を OWIN ミドルウェアが取得する場所になります。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-311">The URL that is passed to the method is where the OWIN middleware obtains instructions for getting the key it needs to verify the signature on the bootstrap token received from the Office host.</span></span> <span data-ttu-id="e5c6e-312">URL の権威セグメントは、web.config から取得されます。これは「common」という文字列か、シングルテナント アドインの場合は GUID です。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-312">The Authority segment of the URL comes from the web.config. It is either the string "common" or, for a single-tenant add-in, a GUID.</span></span>

    ```csharp
    string[] endAuthoritySegments = { "oauth2/v2.0" };
    string[] parsedAuthority = ConfigurationManager.AppSettings["ida:Authority"].Split(endAuthoritySegments, System.StringSplitOptions.None);
    string wellKnownURL = parsedAuthority[0] + "v2.0/.well-known/openid-configuration";

    app.UseOAuthBearerAuthentication(new OAuthBearerAuthenticationOptions
    {
        AccessTokenFormat = new JwtFormat(tvps, new OpenIdConnectCachingSecurityTokenProvider(wellKnownURL))
    });
    ```

1. <span data-ttu-id="e5c6e-313">ファイルを保存して閉じます。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-313">Save and close the file.</span></span>

### <a name="create-the-apivalues-controller"></a><span data-ttu-id="e5c6e-314">/api/values コントローラーを作成する</span><span class="sxs-lookup"><span data-stu-id="e5c6e-314">Create the /api/values controller</span></span>

1. <span data-ttu-id="e5c6e-315">ファイル **Controllers\ValueController.cs** を開きます。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-315">Open the file **Controllers\ValueController.cs**.</span></span> <span data-ttu-id="e5c6e-316">このコントローラーは、SSO システムがブートストラップ トークンを正常に取得した場合に使用されます。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-316">This controller is used when the SSO system has successfully obtained a bootstrap token.</span></span> <span data-ttu-id="e5c6e-317">フォールバック認証システムの一部として使用されることはありません。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-317">It is not used as part of the fallback authorization system.</span></span> <span data-ttu-id="e5c6e-318">そのシステムで AzureADAuthController が使用されました。これは、自動的に作成されます。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-318">That system used the AzureADAuthController, which has been created for you.</span></span>

1. <span data-ttu-id="e5c6e-319">ファイルの先頭に、次に示す `using` ステートメントがあることを確認します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-319">Ensure that the following `using` statements are at the top of the file.</span></span>

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

1. <span data-ttu-id="e5c6e-320">Just above the line that declares the `ValuesController`, add the `[Authorize]` attribute.</span><span class="sxs-lookup"><span data-stu-id="e5c6e-320">Just above the line that declares the `ValuesController`, add the `[Authorize]` attribute.</span></span> <span data-ttu-id="e5c6e-321">This ensures that your add-in will run the authorization process that you configured in the last procedure whenever a controller method is called.</span><span class="sxs-lookup"><span data-stu-id="e5c6e-321">This ensures that your add-in will run the authorization process that you configured in the last procedure whenever a controller method is called.</span></span> <span data-ttu-id="e5c6e-322">Only callers with a valid access token to your add-in can invoke the methods of the controller.</span><span class="sxs-lookup"><span data-stu-id="e5c6e-322">Only callers with a valid access token to your add-in can invoke the methods of the controller.</span></span>

1. <span data-ttu-id="e5c6e-323">次のメソッドを `ValuesController` に追加します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-323">Add the following method to the `ValuesController`.</span></span> <span data-ttu-id="e5c6e-324">戻り値は、`Task<IEnumerable<string>>` ではなく `GET api/values` メソッドでより一般的な `Task<HttpResponseMessage>` になる点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-324">Note that the return value is `Task<HttpResponseMessage>` instead of `Task<IEnumerable<string>>` as would be more common for a `GET api/values` method.</span></span> <span data-ttu-id="e5c6e-325">これは、OAuth 認証論理が ASP.NET フィルターではなく、コントローラーに存在する必要があるということの副作用です。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-325">This is a side effect of that fact that the OAuth  authorization logic must be in the controller, instead of in an ASP.NET filter.</span></span> <span data-ttu-id="e5c6e-326">その論理の一部のエラーの条件では、アドインのクライアントに HTTP 応答オブジェクトが送信される必要があります。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-326">Some error conditions in that logic require that an HTTP Response object be sent to the add-in's client.</span></span>

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

1. <span data-ttu-id="e5c6e-327">`TODO1` を次のコードに置き換えて、`access_as_user` を含むトークンで指定されているスコープを検証します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-327">Replace `TODO1` with the following code to validate that the scopes that are specified in the token include `access_as_user`.</span></span> <span data-ttu-id="e5c6e-328">`SendErrorToClient` メソッドの第 2 パラメーターは、**Exception** オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-328">Note that the second parameter of the `SendErrorToClient` method is an **Exception** object.</span></span> <span data-ttu-id="e5c6e-329">この場合、コードは `null` を渡します。これは、**Exception** オブジェクトが含まれていることで、生成される HTTP 応答には **Message** プロパティが含められなくなるためです。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-329">In this case, the code passes `null` because including the **Exception** object blocks the inclusion of the **Message** property in the HTTP Response that is generated.</span></span>


    ```csharp
    string[] addinScopes = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/scope").Value.Split(' ');
    if (!(addinScopes.Contains("access_as_user")))
    {
        return HttpErrorHelper.SendErrorToClient(HttpStatusCode.Unauthorized, null, "Missing access_as_user.");
    }
    ```

1. <span data-ttu-id="e5c6e-330">`TODO 2` を次のコードに置き換えて、「代理」フローを使用して Microsoft Graph のトークンを取得するために必要なすべての情報を編成します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-330">Replace `TODO 2` with the following code to assemble all the information that is needed to get a token for Microsoft Graph using the "on behalf of" flow.</span></span> <span data-ttu-id="e5c6e-331">このコードについては、以下の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-331">About this code, note:</span></span>

    * <span data-ttu-id="e5c6e-332">Your add-in is no longer playing the role of a resource (or audience) to which the Office host and user need access.</span><span class="sxs-lookup"><span data-stu-id="e5c6e-332">Your add-in is no longer playing the role of a resource (or audience) to which the Office host and user need access.</span></span> <span data-ttu-id="e5c6e-333">Now it is itself a client that needs access to Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="e5c6e-333">Now it is itself a client that needs access to Microsoft Graph.</span></span> <span data-ttu-id="e5c6e-334">`ConfidentialClientApplication` is the MSAL “client context” object.</span><span class="sxs-lookup"><span data-stu-id="e5c6e-334">`ConfidentialClientApplication` is the MSAL “client context” object.</span></span>
    * <span data-ttu-id="e5c6e-335">MSAL.NET 3.x.x からは、`bootstrapContext` は単なるブートストラップ トークンです。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-335">Beginning with MSAL.NET 3.x.x, the `bootstrapContext` is just the bootstrap token itself.</span></span>
    * <span data-ttu-id="e5c6e-336">権威は、web.config から取得されます。これは「common」という文字列か、シングルテナント アドインの場合は GUID です。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-336">The Authority comes from the web.config. It is either the string "common" or, for a single-tenant add-in, a GUID.</span></span>
    * <span data-ttu-id="e5c6e-337">MSAL requires the `openid` and `offline_access` scopes to function, but it throws an error if your code redundantly requests them.</span><span class="sxs-lookup"><span data-stu-id="e5c6e-337">MSAL requires the `openid` and `offline_access` scopes to function, but it throws an error if your code redundantly requests them.</span></span> <span data-ttu-id="e5c6e-338">It will also throw an error if your code requests `profile`, which is really only used when the Office host application gets the token to your add-in's web application.</span><span class="sxs-lookup"><span data-stu-id="e5c6e-338">It will also throw an error if your code requests `profile`, which is really only used when the Office host application gets the token to your add-in's web application.</span></span> <span data-ttu-id="e5c6e-339">So only `Files.Read.All` is explicitly requested.</span><span class="sxs-lookup"><span data-stu-id="e5c6e-339">So only `Files.Read.All` is explicitly requested.</span></span>

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

1. <span data-ttu-id="e5c6e-340">Replace `TODO 3` with the following code.</span><span class="sxs-lookup"><span data-stu-id="e5c6e-340">Replace `TODO 3` with the following code.</span></span> <span data-ttu-id="e5c6e-341">Note about this code:</span><span class="sxs-lookup"><span data-stu-id="e5c6e-341">Note about this code:</span></span>

    * <span data-ttu-id="e5c6e-342">`ConfidentialClientApplication.AcquireTokenOnBehalfOfAsync` メソッドは、最初にメモリ内の MSAL キャッシュで一致するアクセス トークンを探します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-342">The `ConfidentialClientApplication.AcquireTokenOnBehalfOfAsync` method will first look in the MSAL cache, which is in memory, for a matching access token.</span></span> <span data-ttu-id="e5c6e-343">それが見つからなかった場合にのみ、Azure AD V2 エンドポイントで代理フローを開始します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-343">Only if there isn't one, does it initiate the on-behalf-of flow with the Azure AD V2 endpoint.</span></span>
    * <span data-ttu-id="e5c6e-344">`MsalServiceException` 以外の種類の例外は、意図的にキャッチしていないため、`500 Server Error` メッセージとしてクライアントに伝達されます。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-344">Any exceptions that are not of type `MsalServiceException` are intentionally not caught, so they will propagate to the client as `500 Server Error` messages.</span></span>

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

1. <span data-ttu-id="e5c6e-345">`TODO 3a`を以下のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-345">Replace `TODO 3a` with the following code.</span></span> <span data-ttu-id="e5c6e-346">このコードについては、以下の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-346">About this code, note:</span></span>

    * <span data-ttu-id="e5c6e-347">Microsoft Graph リソースが多要素認証を必要としているときに、その認証をユーザーがまだ指定していない場合、Azure AD はエラー `AADSTS50076` と **Claims** プロパティを含む「400 要求が正しくありません」を返します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-347">If multi-factor authentication is required by the Microsoft Graph resource and the user has not yet provided it, Azure AD will return "400 Bad Request" with error `AADSTS50076` and a **Claims** property.</span></span> <span data-ttu-id="e5c6e-348">MSAL は、この情報と共に **MsalUiRequiredException** (**MsalServiceException** から継承) をスローします。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-348">MSAL throws a **MsalUiRequiredException** (which inherits from **MsalServiceException**) with this information.</span></span>
    * <span data-ttu-id="e5c6e-349">**Claims** プロパティの値は、クライアントに渡す必要があります。クライアントは、その値を Office ホストに渡します。Office ホストは、その値を新しいブートストラップ トークンの要求に含めます。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-349">The **Claims** property value must be passed to the client which should pass it to the Office host, which then includes it in a request for a new bootstrap token.</span></span> <span data-ttu-id="e5c6e-350">Azure AD は、認証のすべての要求されたフォームをユーザーに示します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-350">Azure AD will prompt the user for all required forms of authentication.</span></span>
    * <span data-ttu-id="e5c6e-351">The APIs that create HTTP Responses from exceptions don't know about the **Claims** property, so they don't include it in the response object.</span><span class="sxs-lookup"><span data-stu-id="e5c6e-351">The APIs that create HTTP Responses from exceptions don't know about the **Claims** property, so they don't include it in the response object.</span></span> <span data-ttu-id="e5c6e-352">We have to manually create a message that includes it.</span><span class="sxs-lookup"><span data-stu-id="e5c6e-352">We have to manually create a message that includes it.</span></span> <span data-ttu-id="e5c6e-353">A custom **Message** property, however, blocks the creation of an **ExceptionMessage** property, so the only way to get the error ID `AADSTS50076` to the client is to add it to the custom **Message**.</span><span class="sxs-lookup"><span data-stu-id="e5c6e-353">A custom **Message** property, however, blocks the creation of an **ExceptionMessage** property, so the only way to get the error ID `AADSTS50076` to the client is to add it to the custom **Message**.</span></span> <span data-ttu-id="e5c6e-354">JavaScript in the client will need to discover if a response has a **Message** or **ExceptionMessage**, so it knows which to read.</span><span class="sxs-lookup"><span data-stu-id="e5c6e-354">JavaScript in the client will need to discover if a response has a **Message** or **ExceptionMessage**, so it knows which to read.</span></span>
    * <span data-ttu-id="e5c6e-355">カスタム メッセージは、JSON として書式設定されているため、クライアント側の JavaScript は既知の JavaScript `JSON` オブジェクトのメソッドでメッセージを解析できます。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-355">The custom message is formatted as JSON so that the client-side JavaScript can parse it with well-known JavaScript `JSON` object methods.</span></span>

    ```csharp
    if (e.Message.StartsWith("AADSTS50076"))
    {
        string responseMessage = String.Format("{{\"AADError\":\"AADSTS50076\",\"Claims\":{0}}}", e.Claims);
        return HttpErrorHelper.SendErrorToClient(HttpStatusCode.Forbidden, null, responseMessage);
    }
    ```

1. <span data-ttu-id="e5c6e-356">`TODO 3b`を以下のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-356">Replace `TODO 3b` with the following code.</span></span> <span data-ttu-id="e5c6e-357">このコードについては、以下の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-357">About this code, note:</span></span>

    * <span data-ttu-id="e5c6e-358">Azure AD の呼び出しにユーザーまたはテナント管理者のどちらも同意していない (または同意が取り消された) スコープ (アクセス許可) が少なくとも 1 つ含まれていると、Azure AD はエラー `AADSTS65001` と共に「400 要求が正しくありません」を返します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-358">If the call to Azure AD contained at least one scope (permission) for which neither the user nor a tenant administrator has consented (or consent was revoked), Azure AD will return "400 Bad Request" with error `AADSTS65001`.</span></span> <span data-ttu-id="e5c6e-359">MSAL は、この情報と共に **MsalUiRequiredException** をスローします。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-359">MSAL throws a **MsalUiRequiredException** with this information.</span></span>
    *  <span data-ttu-id="e5c6e-360">Azure AD の呼び出しに Azure AD が認識しないスコープが少なくとも 1 つ含まれていると、AAD はエラー `AADSTS70011` と共に「400 要求が正しくありません」を返します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-360">If the call to Azure AD contained at least one scope that Azure AD does not recognize, AAD returns "400 Bad Request" with error `AADSTS70011`.</span></span> <span data-ttu-id="e5c6e-361">MSAL は、この情報と共に **MsalUiRequiredException** をスローします。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-361">MSAL throws a **MsalUiRequiredException** with this information.</span></span>
    *  <span data-ttu-id="e5c6e-362">すべての説明が含まれている理由は、別の条件で 70011 が返されたときに、このアドインでは無効なスコープの存在を意味する場合のみを処理する必要があるためです。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-362">The entire description is included because 70011 is returned in other conditions and it should only be handled in this add-in when it means that there is an invalid scope.</span></span>
    *  <span data-ttu-id="e5c6e-363">The **MsalUiRequiredException** object is passed to `SendErrorToClient`.</span><span class="sxs-lookup"><span data-stu-id="e5c6e-363">The **MsalUiRequiredException** object is passed to `SendErrorToClient`.</span></span> <span data-ttu-id="e5c6e-364">This ensures that an **ExceptionMessage** property that contains the error information is included in the HTTP Response.</span><span class="sxs-lookup"><span data-stu-id="e5c6e-364">This ensures that an **ExceptionMessage** property that contains the error information is included in the HTTP Response.</span></span>

    ```csharp
    if ((e.Message.StartsWith("AADSTS65001")) || (e.Message.StartsWith("AADSTS70011: The provided value for the input parameter 'scope' is not valid.")))
    {
        return HttpErrorHelper.SendErrorToClient(HttpStatusCode.Forbidden, e, null);
    }
    ```

1. <span data-ttu-id="e5c6e-365">`TODO 3c` を次のコードに置き換えて、他のすべての **MsalServiceException** を処理します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-365">Replace `TODO 3c` with the following code to handle all other **MsalServiceException**s.</span></span> <span data-ttu-id="e5c6e-366">前に説明したように、</span><span class="sxs-lookup"><span data-stu-id="e5c6e-366">As noted earlier,</span></span>

    ```csharp
    else
    {
        throw e;
    }
    ```

1. <span data-ttu-id="e5c6e-367">`TODO 4` を次のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-367">Replace `TODO 4` with the following code.</span></span> <span data-ttu-id="e5c6e-368">`GraphApiHelper.GetOneDriveFileNames` メソッドは、自動的に作成されます。これは、Microsoft Graph にデータを要求し、アクセス トークンを含めます。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-368">The `GraphApiHelper.GetOneDriveFileNames` method, which has been created for you, makes the request for data to Microsoft Graph and includes the access token.</span></span>

    ```csharp
    return await GraphApiHelper.GetOneDriveFileNames(authResult.AccessToken);
    ```

1. <span data-ttu-id="e5c6e-369">ファイルを保存して閉じます。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-369">Save and close the file.</span></span>

## <a name="run-the-solution"></a><span data-ttu-id="e5c6e-370">ソリューションを実行する</span><span class="sxs-lookup"><span data-stu-id="e5c6e-370">Run the solution</span></span>

1. <span data-ttu-id="e5c6e-371">Visual Studio ソリューション ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-371">Open the Visual Studio solution file.</span></span>
1. <span data-ttu-id="e5c6e-372">[**ビルド**] メニューで [**ソリューションのクリーン**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-372">On the **Build** menu, select **Clean Solution**.</span></span> <span data-ttu-id="e5c6e-373">終了したら、[**ビルド**] メニューをもう一度開き、[**ソリューションのビルド**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-373">When it finishes, open the **Build** menu again and select **Build Solution**.</span></span>
1. <span data-ttu-id="e5c6e-374">[**ソリューション エクスプローラー**] で、[**Office-Add-in-ASPNET-SSO**] を選択します (一番上のソリューション ノードではなく、「WebAPI」で終わる名前のプロジェクトではありません)。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-374">In **Solution Explorer**, select the **Office-Add-in-ASPNET-SSO** project node (not the top solution node and not the project whose name ends in "WebAPI").</span></span>
1. <span data-ttu-id="e5c6e-375">[**プロパティ**] ウィンドウで、[**ドキュメントの開始**] ドロップダウンを開き、3 つのオプション (Excel、Word、または PowerPoint) のいずれかを選択します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-375">In the **Properties** pane, open the **Start Document** drop down and choose one of the three options (Excel, Word, or PowerPoint).</span></span>

    ![必要な Office ホスト アプリケーション (Excel、PowerPoint、または Word) を選択する](../images/SelectHost.JPG)

1. <span data-ttu-id="e5c6e-377">F5 キーを押します。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-377">Press F5.</span></span>
1. <span data-ttu-id="e5c6e-378">Office アプリケーションの [**ホーム**] リボンで、[**SSO ASP.NET**] グループの [**アドインの表示**] を選択して、タスク ウィンドウ アドインを開きます。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-378">In the Office application, on the **Home** ribbon, select the **Show Add-in** in the **SSO ASP.NET** group to open the task pane add-in.</span></span>
1. <span data-ttu-id="e5c6e-379">[**OneDrive ファイル名の取得**] ボタンをクリックします。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-379">Click the **Get OneDrive File Names** button.</span></span> <span data-ttu-id="e5c6e-380">Microsoft 365 の教育機関または職場のアカウントまたは Microsoft アカウントのいずれかを使用して Office にログインしており、SSO が正常に機能している場合は、OneDrive for Business の最初の10個のファイルとフォルダーの名前が作業ウィンドウに表示されます。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-380">If you are logged into Office with either a Microsoft 365 Education or work account or Microsoft Account, and SSO is working as expected, the first 10 file and folder names in your OneDrive for Business are displayed on the task pane.</span></span> <span data-ttu-id="e5c6e-381">ログインしていない、または SSO をサポートしていないシナリオにいる場合、もしくは何らかの理由で SSO が機能していない場合には、ログインするように求められます。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-381">If you are not logged in, or you are in a scenario that does not support SSO, or SSO is not working for any reason, you will be prompted to log in.</span></span> <span data-ttu-id="e5c6e-382">ログインすると、ファイル名およびフォルダー名が表示されます。</span><span class="sxs-lookup"><span data-stu-id="e5c6e-382">After you log in, the file and folder names appear.</span></span>
