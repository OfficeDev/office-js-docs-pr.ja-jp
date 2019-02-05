---
title: シングル サインオンを使用する ASP.NET Office アドインを作成する
description: ''
ms.date: 01/23/2018
localization_priority: Priority
ms.openlocfilehash: d79dde68b8222a8aafa01a01dc21a4f932f101a9
ms.sourcegitcommit: bf5c56d9b8c573e42bf2268e10ca3fd4d2bb4ff9
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/01/2019
ms.locfileid: "29701821"
---
# <a name="create-an-aspnet-office-add-in-that-uses-single-sign-on-preview"></a><span data-ttu-id="8af07-102">シングル サインオンを使用する ASP.NET Office アドインを作成する (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="8af07-102">Create an ASP.NET Office Add-in that uses single sign-on (preview)</span></span>

<span data-ttu-id="8af07-p101">ユーザーが Office にサインインしたとき、アドインは同じ資格情報を使用し、再度のサインインを要求することなく、複数のアプリケーションへのアクセスを許可することができます。概要については、「[Office アドインで SSO を有効化する](sso-in-office-add-ins.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="8af07-p101">When users are signed in to Office, your add-in can use the same credentials to permit users to access multiple applications without requiring them to sign in a second time. For an overview, see [Enable SSO in an Office Add-in](sso-in-office-add-ins.md).</span></span>

<span data-ttu-id="8af07-105">この記事では、.NET 対応の ASP.NET、OWIN、および Microsoft 認証ライブラリ (MSAL) を使用して作成したアドインで、シングル サインオン (SSO) を有効化するプロセスについて手順を追って説明します。</span><span class="sxs-lookup"><span data-stu-id="8af07-105">This article walks you through the process of enabling single sign-on (SSO) in an add-in that is built with ASP.NET, OWIN, and Microsoft Authentication Library (MSAL) for .NET.</span></span>

> [!NOTE]
> <span data-ttu-id="8af07-106">Node.js ベースのアドインに関する同様の記事については、「[シングル サインオンを使用する Node.js Office アドインを作成する](create-sso-office-add-ins-nodejs.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="8af07-106">For a similar article about a Node.js-based add-in, see [Create a Node.js Office Add-in that uses single sign-on](create-sso-office-add-ins-nodejs.md).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="8af07-107">前提条件</span><span class="sxs-lookup"><span data-stu-id="8af07-107">Prerequisites</span></span>

* <span data-ttu-id="8af07-108">入手可能な Visual Studio 2017 の最新バージョン。</span><span class="sxs-lookup"><span data-stu-id="8af07-108">The latest available version of Visual Studio 2017 Preview.</span></span>

* <span data-ttu-id="8af07-109">Office 365 (「クイック実行」と呼ばれることもあるサブスクリプション バージョン)。</span><span class="sxs-lookup"><span data-stu-id="8af07-109">Office 365 (subscription version, also called “Click to Run”).</span></span> <span data-ttu-id="8af07-110">Insider チャネルからの最新の月次バージョンとビルド。</span><span class="sxs-lookup"><span data-stu-id="8af07-110">Latest monthly version and build from the Insiders channel.</span></span> <span data-ttu-id="8af07-111">このバージョンを入手するには、Office Insider への参加が必要です。</span><span class="sxs-lookup"><span data-stu-id="8af07-111">You might need to be an Office Insider to get this version.</span></span> <span data-ttu-id="8af07-112">詳細については、「[Office Insider になる](https://products.office.com/office-insider?tab=tab-1)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="8af07-112">For more information, see [Be an Office Insider](https://products.office.com/office-insider?tab=tab-1).</span></span> <span data-ttu-id="8af07-113">ビルドが半期チャネルの運用に移行すると、そのビルドで SSO を含むプレビュー機能のサポートはオフになりますので、ご注意ください。</span><span class="sxs-lookup"><span data-stu-id="8af07-113">Please note that when a build graduates to the production semi-annual channel, support for preview features, including SSO, is turned off for that build.</span></span>

## <a name="set-up-the-starter-project"></a><span data-ttu-id="8af07-114">スタート プロジェクトをセットアップする</span><span class="sxs-lookup"><span data-stu-id="8af07-114">Set up the starter project</span></span>

1. <span data-ttu-id="8af07-115">「[Office Add-in ASPNET SSO](https://github.com/officedev/office-add-in-aspnet-sso)」にあるリポジトリを複製するかダウンロードします。</span><span class="sxs-lookup"><span data-stu-id="8af07-115">Clone or download the repo at [Office Add-in ASPNET SSO](https://github.com/officedev/office-add-in-aspnet-sso).</span></span>

1. <span data-ttu-id="8af07-p103">**[Before]** フォルダーを開いて、Visual Studio で .sln ファイルを開きます。これがスタート プロジェクトになります。SSO や承認に直接関連しない UI などの側面は、既に完了しています。</span><span class="sxs-lookup"><span data-stu-id="8af07-p103">Open the **Before** folder and open the .sln file in Visual Studio. This is a starter project. The UI and other aspects of the add-in that are not directly connected to SSO or authorization are already done.</span></span>

    > [!NOTE]
    > <span data-ttu-id="8af07-p104">同じリポジトリ内には、サンプルの完成版も含まれています。これは、この記事の手順を完了したときに得られるアドインと同様のものですが、完成済みのプロジェクトには、この記事のテキストと重複するコード コメントが含まれています。完成版を使用する場合は、`sln` ファイルを開いて、この記事の手順をそのまま実行しますが、「**クライアント側のコードを作成する**」と「**サーバー側のコードを作成する**」のセクションは省略してください。</span><span class="sxs-lookup"><span data-stu-id="8af07-p104">There is also a completed version of the sample in the same repo. It is just like the add-in that you would have if you completed the procedures in this article, except that the completed project has code comments that would be redundant with the text of this article. To use the completed version, just open the `sln` file and follow the instructions in this article, but skip the sections **Code the client side** and **Code the server** side.</span></span>

1. <span data-ttu-id="8af07-p105">プロジェクトを開いたら、そのプロジェクトを Visual Studio でビルドします。その結果として、packages.config ファイルにリストされたパッケージがインストールされます。コンピューターのローカル パッケージ キャッシュに含まれるパッケージの数に応じて、数秒から数分の時間がかかります。</span><span class="sxs-lookup"><span data-stu-id="8af07-p105">After the project opens, build it in Visual Studio, which will install the packages listed in the packages.config file. This can take a few seconds to several minutes depending on how many of the packages are in the computer's local package cache.</span></span>

    > [!NOTE]
    > <span data-ttu-id="8af07-124">ID 名前空間に関するエラーが表示されます。</span><span class="sxs-lookup"><span data-stu-id="8af07-124">You will get an error about the Identity namespace.</span></span> <span data-ttu-id="8af07-125">これは構成の問題の副作用ですが、次のステップで修正します。</span><span class="sxs-lookup"><span data-stu-id="8af07-125">This is a side effect of a configuration issue that you will fix with the next step.</span></span> <span data-ttu-id="8af07-126">重要な点は、パッケージがインストールされていることです。</span><span class="sxs-lookup"><span data-stu-id="8af07-126">The important thing is that the packages are installed.</span></span>

1. <span data-ttu-id="8af07-127">現在、SSO (バージョン `1.1.4-preview0002`) に必要な MSAL ライブラリ (Microsoft.Identity.Client) は標準の nuget カタログの一部ではないため、package.config にはリストされていません。これは、個別にインストールする必要があります。</span><span class="sxs-lookup"><span data-stu-id="8af07-127">Currently, the version of the MSAL library (Microsoft.Identity.Client) that you need for SSO (version `1.1.4-preview0002`) is not part of the standard nuget catalog, so it is not listed in the package.config, and it must be installed separately.</span></span>

   > 1. <span data-ttu-id="8af07-128">**[ツール]** メニューで **[Nuget パッケージ マネージャー]** > **[パッケージ マネージャー コンソール]** に移動します。</span><span class="sxs-lookup"><span data-stu-id="8af07-128">On the **Tools** menu, navigate to **Nuget Package Manager** > **Package Manager Console**.</span></span> 

   > 2. <span data-ttu-id="8af07-129">コンソールで、次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="8af07-129">At the console, run the following command.</span></span> <span data-ttu-id="8af07-130">これは高速インターネット接続の場合でも、完了までに数分かかることがあります。</span><span class="sxs-lookup"><span data-stu-id="8af07-130">It may take a minute or more to complete even with a fast Internet connection.</span></span> <span data-ttu-id="8af07-131">完了すると、コンソールの出力の末尾近くに "**'Microsoft.Identity.Client 1.1.4-preview0002' が正常にインストールされました**" というメッセージが表示されます。</span><span class="sxs-lookup"><span data-stu-id="8af07-131">When it finishes you should see **Successfully installed 'Microsoft.Identity.Client 1.1.4-preview0002' ...** near the end of the output in the console.</span></span>

   >    `Install-Package Microsoft.Identity.Client -Version 1.1.4-preview0002`

   > 3. <span data-ttu-id="8af07-132">**ソリューション エクスプローラー**で、[**Office-Add-in-ASPNET-SSO-WebAPI**] プロジェクトの [**参照設定**] を展開します。</span><span class="sxs-lookup"><span data-stu-id="8af07-132">In **Solution Explorer**, expand **References** of **Office-Add-in-ASPNET-SSO-WebAPI** project.</span></span> <span data-ttu-id="8af07-133">[**Microsoft.Identity.Client**] がリストされていることを確認します。</span><span class="sxs-lookup"><span data-stu-id="8af07-133">Verify that **Microsoft.Identity.Client** is listed.</span></span> <span data-ttu-id="8af07-134">リストされていない場合やそのエントリに警告アイコンが表示されている場合は、エントリを削除してから Visual Studio の [参照の追加] ウィザードを使用し、**... \[Begin | Complete]\packages\Microsoft.Identity.Client.1.1.4-preview0002\lib\net45\Microsoft.Identity.Client.dll** のアセンブリに参照を追加します。</span><span class="sxs-lookup"><span data-stu-id="8af07-134">If it is not or there is a warning icon on its entry, delete the entry and then use the Visual Studio Add Reference Wizard to add a reference to the assembly at **... \[Begin | Complete]\packages\Microsoft.Identity.Client.1.1.4-preview0002\lib\net45\Microsoft.Identity.Client.dll**</span></span>

1. <span data-ttu-id="8af07-135">もう一度プロジェクトをビルドします。</span><span class="sxs-lookup"><span data-stu-id="8af07-135">Build the project a second time.</span></span>

## <a name="register-the-add-in-with-azure-ad-v20-endpoint"></a><span data-ttu-id="8af07-136">Azure AD v2.0 エンドポイントにアドインを登録する</span><span class="sxs-lookup"><span data-stu-id="8af07-136">Register the add-in with Azure AD v2.0 endpoint</span></span>

<span data-ttu-id="8af07-137">次の手順は、複数の場所で使用できるように、一般的に記述されています。</span><span class="sxs-lookup"><span data-stu-id="8af07-137">The following instruction are written generically so they can be used in multiple places.</span></span> <span data-ttu-id="8af07-138">この記事では、次の手順を行います。</span><span class="sxs-lookup"><span data-stu-id="8af07-138">For this ariticle do the following:</span></span>
- <span data-ttu-id="8af07-139">プレースホルダー **$ADD-IN-NAME$** を `Office-Add-in-ASPNET-SSO` に置き換えます。</span><span class="sxs-lookup"><span data-stu-id="8af07-139">Replace the placeholder **$ADD-IN-NAME$** with `Office-Add-in-ASPNET-SSO`.</span></span>
- <span data-ttu-id="8af07-140">プレースホルダー **$FQDN-WITHOUT-PROTOCOL$** を `localhost:44355` に置き換えます。</span><span class="sxs-lookup"><span data-stu-id="8af07-140">Replace the placeholder **$FQDN-WITHOUT-PROTOCOL$** with `localhost:44355`.</span></span>
- <span data-ttu-id="8af07-141">[**アクセス許可の選択**] ダイアログでアクセス許可を指定するときに、次のアクセス許可のチェック ボックスをオンにします。</span><span class="sxs-lookup"><span data-stu-id="8af07-141">When you specify permissions in the **Select Permissions** dialog, check the boxes for the following permissions.</span></span> <span data-ttu-id="8af07-142">実際にアドイン自体に必要なのは最初のものだけですが、サーバー側コードで使用される MSAL ライブラリで `offline_access` および `openid` が必要とされます。</span><span class="sxs-lookup"><span data-stu-id="8af07-142">Only the first is really required by your add-in itself; but the MSAL library that the server-side code uses requires `offline_access` and `openid`.</span></span> <span data-ttu-id="8af07-143">Office ホストがアドインの Web アプリケーションに対してトークンを取得するために、`profile` のアクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="8af07-143">The `profile` permission is required for the Office host to get a token to your add-in web application.</span></span>
    * <span data-ttu-id="8af07-144">Files.Read.All</span><span class="sxs-lookup"><span data-stu-id="8af07-144">Files.Read.All</span></span>
    * <span data-ttu-id="8af07-145">offline_access</span><span class="sxs-lookup"><span data-stu-id="8af07-145">offline_access</span></span>
    * <span data-ttu-id="8af07-146">openid</span><span class="sxs-lookup"><span data-stu-id="8af07-146">openid</span></span>
    * <span data-ttu-id="8af07-147">profile</span><span class="sxs-lookup"><span data-stu-id="8af07-147">profile</span></span>


[!INCLUDE[](../includes/register-sso-add-in-aad-v2-include.md)]

## <a name="grant-administrator-consent-to-the-add-in"></a><span data-ttu-id="8af07-148">アドインに管理者の同意を許可する</span><span class="sxs-lookup"><span data-stu-id="8af07-148">Grant administrator consent to the add-in</span></span>

[!INCLUDE[](../includes/grant-admin-consent-to-an-add-in-include.md)]

## <a name="configure-the-add-in"></a><span data-ttu-id="8af07-149">アドインを構成する</span><span class="sxs-lookup"><span data-stu-id="8af07-149">Configure the add-in</span></span>

1. <span data-ttu-id="8af07-150">次の文字列で、プレースホルダー "{tenant_ID}" を Office 365 テナント ID に置き換えます。</span><span class="sxs-lookup"><span data-stu-id="8af07-150">In the following string, replace the placeholder “{tenant_ID}” with your Office 365 tenant ID.</span></span> <span data-ttu-id="8af07-151">「[Office 365 のテナント ID を検索します](https://docs.microsoft.com/onedrive/find-your-office-365-tenant-id)」のいずれかのメソッドを使用して、テナント ID を取得します。</span><span class="sxs-lookup"><span data-stu-id="8af07-151">Use one of the methods in [Find your Office 365 tenant ID](https://docs.microsoft.com/onedrive/find-your-office-365-tenant-id) to obtain it.</span></span>

    `https://login.microsoftonline.com/{tenant_ID}/v2.0`

2. <span data-ttu-id="8af07-152">Visual Studio で、web.config を開きます。**appSettings** セクションには、値を割り当てる必要のあるいくつかのキーがあります。</span><span class="sxs-lookup"><span data-stu-id="8af07-152">In Visual Studio, open the web.config. There are some keys in the **appSettings** section to which you need to assign values.</span></span>

3. <span data-ttu-id="8af07-p112">"ida:Issuer" という名前のキーの値として、手順 1 で作成した文字列を使用します。この値に、空白スペースが含まれていないことを確認してください。</span><span class="sxs-lookup"><span data-stu-id="8af07-p112">Use the string you constructed in step 1 as the value to the key named “ida:Issuer”. Be sure there are no blank spaces in the value.</span></span>

4. <span data-ttu-id="8af07-155">次に示す値を対応するキーに代入します。</span><span class="sxs-lookup"><span data-stu-id="8af07-155">Assign the following values to the corresponding keys:</span></span>

    |<span data-ttu-id="8af07-156">キー</span><span class="sxs-lookup"><span data-stu-id="8af07-156">Key</span></span>|<span data-ttu-id="8af07-157">値</span><span class="sxs-lookup"><span data-stu-id="8af07-157">Value</span></span>|
    |:-----|:-----|
    |<span data-ttu-id="8af07-158">ida:ClientID</span><span class="sxs-lookup"><span data-stu-id="8af07-158">ida:ClientID</span></span>|<span data-ttu-id="8af07-159">アドインの登録時に取得したアプリケーション ID。</span><span class="sxs-lookup"><span data-stu-id="8af07-159">The application ID you obtained when you registered the add-in.</span></span>|
    |<span data-ttu-id="8af07-160">ida:Audience</span><span class="sxs-lookup"><span data-stu-id="8af07-160">ida:Audience</span></span>|<span data-ttu-id="8af07-161">アドインの登録時に取得したアプリケーション ID。</span><span class="sxs-lookup"><span data-stu-id="8af07-161">The application ID you obtained when you registered the add-in.</span></span>|
    |<span data-ttu-id="8af07-162">ida:Password</span><span class="sxs-lookup"><span data-stu-id="8af07-162">ida:Password</span></span>|<span data-ttu-id="8af07-163">アドインの登録時に取得したパスワード。</span><span class="sxs-lookup"><span data-stu-id="8af07-163">The password you obtained when you registered the add-in.</span></span>|

   <span data-ttu-id="8af07-p113">次に、4 つのキーの変更後の例を示します。*ClientID と Audience が同じになっている点に注目してください*。両方の目的に単一のキーを使用することもできますが、これらは必ずしも同じではないため、別々に保持しておくと web.config のマークアップが再利用しやすくなります。また、別のキーを使用することで、アドインが Office ホストに関連する OAuth リソースと、Microsoft Graph に関連する OAuth クライアントの両方でであるという考えが補強されます。</span><span class="sxs-lookup"><span data-stu-id="8af07-p113">The following is an example of what the four keys you changed should look like. *Note that ClientID and Audience are the same*. You can also use a single key for both purposes, but your web.config markup is more reusable if you keep them separate because they aren't always the same. Also, having separate keys reinforces the idea that your add-in is both an OAuth resource, relative to the Office host, and an OAuth client, relative to Microsoft Graph.</span></span>

    ```xml
    <add key=”ida:ClientID" value="12345678-1234-1234-1234-123456789012" />
    <add key="ida:Audience" value="12345678-1234-1234-1234-123456789012" />
    <add key="ida:Password" value="rFfv17ezsoGw5XUc0CDBHiU" />
    <add key="ida:Issuer" value="https://login.microsoftonline.com/aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee/v2.0" />
    
    ```

   > [!NOTE]
   > <span data-ttu-id="8af07-168">その他の **[appSettings]** セクションの設定は、未変更のままにします。</span><span class="sxs-lookup"><span data-stu-id="8af07-168">Leave the other settings in the **appSettings** section unchanged.</span></span>

1. <span data-ttu-id="8af07-169">ファイルを保存して閉じます。</span><span class="sxs-lookup"><span data-stu-id="8af07-169">Save and close the file.</span></span>

1. <span data-ttu-id="8af07-170">アドイン プロジェクトで、アドイン マニフェスト ファイル "Office-Add-in-ASPNET-SSO.xml" を開きます。</span><span class="sxs-lookup"><span data-stu-id="8af07-170">In the add-in project, open the add-in manifest file “Office-Add-in-ASPNET-SSO.xml”.</span></span>

1. <span data-ttu-id="8af07-171">ファイルの最後までスクロールします。</span><span class="sxs-lookup"><span data-stu-id="8af07-171">Scroll to the bottom of the file.</span></span>

1. <span data-ttu-id="8af07-172">`</VersionOverrides>` 終了タグの直前に、以下のマークアップがあります。</span><span class="sxs-lookup"><span data-stu-id="8af07-172">Just above the end `</VersionOverrides>` tag, you'll find the following markup:</span></span>

    ```xml
    <WebApplicationInfo>
      <Id>{application_GUID here}</Id>
      <Resource>api://localhost:44355/{application_GUID here}</Resource>
      <Scopes>
          <Scope>Files.Read.All</Scope>
          <Scope>offline_access</Scope>
          <Scope>openid</Scope>
          <Scope>profile</Scope>
      </Scopes>
    </WebApplicationInfo>
    ```

1. <span data-ttu-id="8af07-173">このマークアップ内の*両方の場所の*プレースホルダー “{application_GUID here}” を、アドインの登録時にコピーしたアプリケーション ID に置き換えます。</span><span class="sxs-lookup"><span data-stu-id="8af07-173">Replace the placeholder “{application_GUID here}” *in both places* in the markup with the Application ID that you copied when you registered your add-in.</span></span> <span data-ttu-id="8af07-174">"{}" は ID の一部ではないので、これらを含めないでください。</span><span class="sxs-lookup"><span data-stu-id="8af07-174">The "{}" are not part of the ID, so do not include them.</span></span> <span data-ttu-id="8af07-175">これは、web.config の ClientID と Audience に使用したものと同じ ID です。</span><span class="sxs-lookup"><span data-stu-id="8af07-175">This is the same ID you used in for the ClientID and Audience in the web.config.</span></span>

    > [!NOTE]
    > * <span data-ttu-id="8af07-176">**リソース**の値は、アドインの登録に Web API プラットフォームを追加したときに設定した**アプリケーション ID URI** です。</span><span class="sxs-lookup"><span data-stu-id="8af07-176">The **Resource** value is the **Application ID URI** you set when you added the Web API platform to the registration of the add-in.</span></span>
    > * <span data-ttu-id="8af07-177">**[範囲]** セクションは、アドインが AppSource から販売された場合に、同意ダイアログ ボックスを生成するためにのみ使用します。</span><span class="sxs-lookup"><span data-stu-id="8af07-177">The **Scopes** section is used only to generate a consent dialog box if the add-in is sold through AppSource.</span></span>

1. <span data-ttu-id="8af07-178">Visual Studio で、**[エラー一覧]** の **[警告]** タブを開きます。</span><span class="sxs-lookup"><span data-stu-id="8af07-178">Open the **Warnings** tab of the **Error List** in Visual Studio.</span></span> <span data-ttu-id="8af07-179">`<WebApplicationInfo>` が `<VersionOverrides>` の有効な子ではないという警告が表示される場合は、Visual Studio 2017 プレビューのバージョンで SSO マークアップが認識されていません。</span><span class="sxs-lookup"><span data-stu-id="8af07-179">If there is a warning that `<WebApplicationInfo>` is not a valid child of `<VersionOverrides>`, your version of Visual Studio 2017 Preview does not recognize the SSO markup.</span></span> <span data-ttu-id="8af07-180">回避策として、Word、Excel、または PowerPoint のアドインに対して、次の操作を行います。</span><span class="sxs-lookup"><span data-stu-id="8af07-180">As a workaround, do the following for a Word, Excel, or PowerPoint add-in.</span></span> <span data-ttu-id="8af07-181">(Outlook アドインを使用している場合は、以下の回避策を参照してください。)</span><span class="sxs-lookup"><span data-stu-id="8af07-181">(If you are working with an Outlook add-in see the workaround below.)</span></span>

   - <span data-ttu-id="8af07-182">**Word、Excel、および PowerPoint の回避策**</span><span class="sxs-lookup"><span data-stu-id="8af07-182">**Workaround for Word, Excel, and Powerpoint**</span></span>

        1. <span data-ttu-id="8af07-183">マニフェストの `</VersionOverrides>` の終了タグの直前の `<WebApplicationInfo>` セクションをコメント アウトします。</span><span class="sxs-lookup"><span data-stu-id="8af07-183">Comment out the `<WebApplicationInfo>` section from the manifest just above the end of `</VersionOverrides>`.</span></span>

        2. <span data-ttu-id="8af07-p116">**F5** キーを押してデバッグ セッションを開始します。これにより、次のフォルダーにマニフェストのコピーが作成されます (これには、Visual Studio よりも**ファイル エクスプローラー**の方が容易にアクセスできます): `Office-Add-in-ASP.NET-SSO\Complete\Office-Add-in-ASPNET-SSO\bin\Debug\OfficeAppManifests`</span><span class="sxs-lookup"><span data-stu-id="8af07-p116">Press **F5** to start a debugging session. This will create a copy of the manifest in the following folder (which is easier to access in **File Explorer** than in Visual Studio): `Office-Add-in-ASP.NET-SSO\Complete\Office-Add-in-ASPNET-SSO\bin\Debug\OfficeAppManifests`</span></span>

        3. <span data-ttu-id="8af07-186">マニフェストのコピーから、`<WebApplicationInfo>` セクションの周囲のコメント構文を削除します。</span><span class="sxs-lookup"><span data-stu-id="8af07-186">In the copy of the manifest, remove the comment syntax around the `<WebApplicationInfo>` section.</span></span>

        4. <span data-ttu-id="8af07-187">マニフェストのコピーを保存します。</span><span class="sxs-lookup"><span data-stu-id="8af07-187">Save the copy of the manifest.</span></span>

        5. <span data-ttu-id="8af07-p117">この時点で、次回 F5 キーを押したときに、このマニフェストのコピーが Visual Studio によって上書きされないようにする必要があります。**ソリューション エクスプローラー**の上部にあるソリューション ノード (どちらのプロジェクト ノードでもない) を右クリックします。</span><span class="sxs-lookup"><span data-stu-id="8af07-p117">Now you must prevent Visual Studio from overwriting the copy of the manifest the next time you press F5. Right-click the solution node at the very top of **Solution Explorer** (not either of the project nodes).</span></span>

        6. <span data-ttu-id="8af07-190">コンテキスト メニューから **[プロパティ]** を選択します。**[ソリューション プロパティ ページ]** ダイアログ ボックスが開きます。</span><span class="sxs-lookup"><span data-stu-id="8af07-190">Select **Properties** from the context menu and a **Solution Property Pages** dialog box opens.</span></span>

        7. <span data-ttu-id="8af07-191">**[構成プロパティ]** を展開し、**[構成]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="8af07-191">Expand **Configuration Properties** and select **Configuration**.</span></span>

        8. <span data-ttu-id="8af07-192">**Office-Add-in-ASPNET-SSO** プロジェクト (**Office-Add-in-ASPNET-SSO-WebAPI** プロジェクトでは*ありません*) の行で、**[ビルド]** と **[展開]** を選択解除します。</span><span class="sxs-lookup"><span data-stu-id="8af07-192">Deselect **Build** and **Deploy** in the row for the **Office-Add-in-ASPNET-SSO** project (*not* the **Office-Add-in-ASPNET-SSO-WebAPI** project).</span></span>

        9. <span data-ttu-id="8af07-193">**[OK]** をクリックしてダイアログ ボックスを閉じます。</span><span class="sxs-lookup"><span data-stu-id="8af07-193">Press **OK** to close the dialog box.</span></span>

   - <span data-ttu-id="8af07-194">**Outlook の回避策**</span><span class="sxs-lookup"><span data-stu-id="8af07-194">**Workaround for Outlook**</span></span>

        1. <span data-ttu-id="8af07-195">開発用コンピューターで、既存の `MailAppVersionOverridesV1_1.xsd` を探します。</span><span class="sxs-lookup"><span data-stu-id="8af07-195">On your development machine, locate the existing `MailAppVersionOverridesV1_1.xsd`.</span></span> <span data-ttu-id="8af07-196">`./Xml/Schemas/{lcid}` の下の Visual Studio インストール ディレクトリに配置されています。</span><span class="sxs-lookup"><span data-stu-id="8af07-196">This should be located in your Visual Studio installation directory under `./Xml/Schemas/{lcid}`.</span></span> <span data-ttu-id="8af07-197">たとえば、英語版 (米国) の VS 2017 32 ビットの標準インストールの場合、完全なパスは、`C:\Program Files (x86)\Microsoft Visual Studio\2017\Enterprise\Xml\Schemas\1033` になります。</span><span class="sxs-lookup"><span data-stu-id="8af07-197">For example, on a typical installation of VS 2017 32-bit on an English (US) system, the full path would be `C:\Program Files (x86)\Microsoft Visual Studio\2017\Enterprise\Xml\Schemas\1033`.</span></span>

        2. <span data-ttu-id="8af07-198">既存のファイルの名前を、`MailAppVersionOverridesV1_1.old` に変更します。</span><span class="sxs-lookup"><span data-stu-id="8af07-198">Rename the existing file to `MailAppVersionOverridesV1_1.old`.</span></span>

        3. <span data-ttu-id="8af07-199">変更したこのファイルを、フォルダーにコピーします。[変更済みの MailAppVersionOverrides スキーマ](https://github.com/OfficeDev/outlook-add-in-attachments-demo/blob/sso-conversion/manifest-schema-fix/MailAppVersionOverridesV1_1.xsd)</span><span class="sxs-lookup"><span data-stu-id="8af07-199">Copy this modified version of the file into the folder: [Modified MailAppVersionOverrides Schema](https://github.com/OfficeDev/outlook-add-in-attachments-demo/blob/sso-conversion/manifest-schema-fix/MailAppVersionOverridesV1_1.xsd)</span></span>

1. <span data-ttu-id="8af07-200">Visual Studio でメインのマニフェスト ファイルを保存して閉じます。</span><span class="sxs-lookup"><span data-stu-id="8af07-200">Save and close the main manifest file in Visual Studio.</span></span>

## <a name="code-the-client-side"></a><span data-ttu-id="8af07-201">クライアント側のコードの作成</span><span class="sxs-lookup"><span data-stu-id="8af07-201">Code the client side</span></span>

1. <span data-ttu-id="8af07-p119">**[Scripts]** フォルダー内の Home.js ファイルを開きます。これには、一部のコードが既に含まれています。</span><span class="sxs-lookup"><span data-stu-id="8af07-p119">Open the Home.js file in the **Scripts** folder. It already has some code in it:</span></span>
    * <span data-ttu-id="8af07-204">`Office.initialize` メソッドへの割り当てが、`getGraphAccessTokenButton` ボタン クリック イベントへのハンドラーの割り当てになります。</span><span class="sxs-lookup"><span data-stu-id="8af07-204">An assignment to the `Office.initialize` method that, in turn, assigns a handler to the `getGraphAccessTokenButton` button click event.</span></span>
    * <span data-ttu-id="8af07-205">`showResult` メソッドは、作業ウィンドウの下側に Microsoft Graph から返されたデータ (またはエラー メッセージ) を表示するものです。</span><span class="sxs-lookup"><span data-stu-id="8af07-205">A `showResult` method that will display data returned from Microsoft Graph (or an error message) at the bottom of the task pane.</span></span>
    * <span data-ttu-id="8af07-206">`logErrors` メソッドは、エンド ユーザーを対象としていないエラーをコンソールにログ出力するものです。</span><span class="sxs-lookup"><span data-stu-id="8af07-206">A `logErrors` method that will log to console errors that are not intended for the end user.</span></span>

1. <span data-ttu-id="8af07-p120">`Office.initialize` への割り当ての下に、次に示すコードを追加します。このコードについては、次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="8af07-p120">Below the assignment to `Office.initialize`, add the code below. Note the following about this code:</span></span>

    * <span data-ttu-id="8af07-209">アドインのエラー処理により、アクセス トークンの取得が別のオプションのセットを使用して自動的に再試行されることがあります。</span><span class="sxs-lookup"><span data-stu-id="8af07-209">The error-handling in the add-in will sometimes automatically attempt a second time to get an access token, using a different set of options.</span></span> <span data-ttu-id="8af07-210">カウンター変数 `timesGetOneDriveFilesHasRun` とフラグ変数 `triedWithoutForceConsent` を使用して、失敗するトークン取得の繰り返しからユーザーが抜け出せるようにします。</span><span class="sxs-lookup"><span data-stu-id="8af07-210">The counter variable `timesGetOneDriveFilesHasRun`, and the flag variable `triedWithoutForceConsent` are used to ensure that the user isn't cycled repeatedly through failed attempts to get a token.</span></span> 
    * <span data-ttu-id="8af07-p122">この後の手順では `getDataWithToken` メソッドを作成しますが、そのメソッドで `forceConsent` というオプションが `false` に設定される点に注意してください。詳細については、次の手順で説明します。</span><span class="sxs-lookup"><span data-stu-id="8af07-p122">You create the `getDataWithToken` method in the next step, but note that it sets an option called `forceConsent` to `false`. More about that in the next step.</span></span>

    ```javascript
    var timesGetOneDriveFilesHasRun = 0;
    var triedWithoutForceConsent = false;

    function getOneDriveFiles() {
        timesGetOneDriveFilesHasRun++;
        triedWithoutForceConsent = true;
        getDataWithToken({ forceConsent: false });
    }   
    ```

1. <span data-ttu-id="8af07-p123">`getOneDriveFiles` メソッドの下に、次のコードを追加します。このコードについては、次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="8af07-p123">Below the `getOneDriveFiles` method, add the code below. Note the following about this code:</span></span>

    * <span data-ttu-id="8af07-215">[getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) は Office.js の新しい API です。これにより、アドインは Office ホスト アプリケーション (Excel、PowerPoint、Word など) に、アドインへの (Office にサインインしているユーザーの) アクセス トークンを要求できるようになります。</span><span class="sxs-lookup"><span data-stu-id="8af07-215">The [getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) is the new API in Office.js that enables an add-in to ask the Office host application (Excel, PowerPoint, Word, etc.) for an access token to the add-in (for the user signed into Office).</span></span> <span data-ttu-id="8af07-216">その結果、この Office ホスト アプリケーションによって、Azure AD 2.0 エンドポイントにこのトークンが要求されます。</span><span class="sxs-lookup"><span data-stu-id="8af07-216">The Office host application, in turn, asks the Azure AD 2.0 endpoint for the token.</span></span> <span data-ttu-id="8af07-217">アドインの登録時に、アドインに対する Office ホストを事前認証しているため、Azure AD はそのトークンを送信します。</span><span class="sxs-lookup"><span data-stu-id="8af07-217">Since you preauthorized the Office host to your add-in when you registered it, Azure AD will send the token.</span></span>
    * <span data-ttu-id="8af07-218">Office にサインインしているユーザーがいない場合、Office ホストはユーザーにサインインを求めるダイアログを表示します。</span><span class="sxs-lookup"><span data-stu-id="8af07-218">If no user is signed into Office, the Office host will prompt the user to sign in.</span></span>
    * <span data-ttu-id="8af07-219">オプションのパラメーター `forceConsent` を `false` に設定すると、ユーザーがアドインを使用するたびに、Office ホストにアドインへのアクセス権を付与するための同意を求めるダイアログが表示されなくなります。</span><span class="sxs-lookup"><span data-stu-id="8af07-219">The options parameter sets `forceConsent` to `false`, so the user will not be prompted to consent to giving the Office host access to your add-in every time she or he uses the add-in.</span></span> <span data-ttu-id="8af07-220">ユーザーが初めてアドインを実行すると、`getAccessTokenAsync` の呼び出しは失敗しますが、この後の手順で追加するエラー処理ロジックにより、`forceConsent` オプションを `true` に設定した再呼び出しが自動的に実行され、ユーザーに同意を求めるダイアログが表示されます。ただし、これは初回時のみ実行されます。</span><span class="sxs-lookup"><span data-stu-id="8af07-220">The first time the user runs the add-in, the call of `getAccessTokenAsync` will fail, but error-handling logic that you add in a later step will automatically re-call with the `forceConsent` option set to `true` and the user will be prompted to consent, but only that first time.</span></span>
    * <span data-ttu-id="8af07-221">`handleClientSideErrors` メソッドは、この後の手順で作成します。</span><span class="sxs-lookup"><span data-stu-id="8af07-221">You will create the `handleClientSideErrors` method in a later step.</span></span>

    ```javascript
    function getDataWithToken(options) {
    Office.context.auth.getAccessTokenAsync(options,
        function (result) {
            if (result.status === "succeeded") {
                TODO1: Use the access token to get Microsoft Graph data.
            }
            else {
                handleClientSideErrors(result);
            }
        });
    }
    ```

1. <span data-ttu-id="8af07-p126">TODO1 を次に示す行に置き換えます。`getData` メソッドとサーバー側の "/api/values" ルートは、この後の手順で作成します。エンドポイントには、相対 URL を使用します。これは、その URL がアドインと同じドメインでホストされている必要があるためです。</span><span class="sxs-lookup"><span data-stu-id="8af07-p126">Replace the TODO1 with the following lines. You create the `getData` method and the server-side “/api/values” route in later steps. A relative URL is used for the endpoint because it must be hosted on the same domain as your add-in.</span></span>

    ```javascript
    accessToken = result.value;
    getData("/api/values", accessToken);
    ```

1. <span data-ttu-id="8af07-p127">`getOneDriveFiles` メソッドの下に、以下を追加します。このコードについては、次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="8af07-p127">Below the `getOneDriveFiles` method, add the following. About this code, note:</span></span>

    * <span data-ttu-id="8af07-p128">このメソッドは、特定の Web API エンドポイントを呼び出して、Office ホスト アプリケーションがアドインへのアクセスに使用したものと同じアクセス トークンを渡します。サーバー側では、このアクセス トークンが Microsoft Graph へのアクセス トークンを取得するための「代理 (on-behalf-of)」フローで使用されます。</span><span class="sxs-lookup"><span data-stu-id="8af07-p128">This method calls a specified Web API endpoint and passes it the same access token that the Office host application used to get access to your add-in. On the server-side, this access token will be used in the “on behalf of” flow to obtain an access token to Microsoft Graph.</span></span>
    * <span data-ttu-id="8af07-229">`handleServerSideErrors` メソッドは、この後の手順で作成します。</span><span class="sxs-lookup"><span data-stu-id="8af07-229">You will create the `handleServerSideErrors` method in a later step.</span></span>

    ```javascript
    function getData(relativeUrl, accessToken) {
        $.ajax({
            url: relativeUrl,
            headers: { "Authorization": "Bearer " + accessToken },
            type: "GET"
        })
        .done(function (result) {
            showResult(result);
        })
        .fail(function (result) {
            handleServerSideErrors(result);
        }); 
    }
    ```

### <a name="create-the-error-handling-methods"></a><span data-ttu-id="8af07-230">エラー処理のメソッドを作成する</span><span class="sxs-lookup"><span data-stu-id="8af07-230">Create the error-handling methods</span></span>

1. <span data-ttu-id="8af07-231">`getData` メソッドの下に、次のメソッドを追加します。</span><span class="sxs-lookup"><span data-stu-id="8af07-231">Below the `getData` method, add the following method.</span></span> <span data-ttu-id="8af07-232">このメソッドは、Office ホストがアドインの Web サービスへのアクセス トークンを取得できないときに、アドインのクライアントでエラーを処理します。</span><span class="sxs-lookup"><span data-stu-id="8af07-232">This method will handle errors in the add-in's client when the Office host is unable to obtain an access token to the add-in's web service.</span></span> <span data-ttu-id="8af07-233">こうしたエラーはエラー コードで報告されるため、このメソッドでは `switch` ステートメントを使用してエラーを識別します。</span><span class="sxs-lookup"><span data-stu-id="8af07-233">These errors are reported with an error code, so the method uses a `switch` statement to distinguish them.</span></span>

    ```javascript
    function handleClientSideErrors(result) {

        switch (result.error.code) {
    
            // TODO2: Handle the case where user is not logged in, or the user cancelled, without responding, a
            //        prompt to provide a 2nd authentication factor. 
    
            // TODO3: Handle the case where the user's sign-in or consent was aborted.
    
            // TODO4: Handle the case where the user is logged in with an account that is neither work or school, 
            //        nor Micrososoft Account.
    
            // TODO5: Handle an unspecified error from the Office host.
    
            // TODO6: Handle the case where the Office host cannot get an access token to the add-ins 
            //        web service/application.
    
            // TODO7: Handle the case where the user tiggered an operation that calls `getAccessTokenAsync` 
            //        before a previous call of it completed.
    
            // TODO8: Handle the case where the add-in does not support forcing consent.
    
            // TODO9: Log all other client errors.
        }
    }
    ```

1. <span data-ttu-id="8af07-234">`TODO2` を次のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="8af07-234">Replace `TODO2` with the following code.</span></span> <span data-ttu-id="8af07-235">エラー 13001 は、ユーザーがログインしていない場合、または 2 番目の認証要素の指定を求めるダイアログに応答しないでキャンセルした場合に発生します。</span><span class="sxs-lookup"><span data-stu-id="8af07-235">Error 13001 occurs when the user is not logged in, or the user cancelled, without responding, a prompt to provide a 2nd authentication factor.</span></span> <span data-ttu-id="8af07-236">どちらの場合も、このコードでは `getDataWithToken` メソッドを再実行して、サインインを求めるダイアログの表示を強制するようにオプションを設定します。</span><span class="sxs-lookup"><span data-stu-id="8af07-236">In either case, the code re-runs the `getDataWithToken` method and sets an option to force a sign-in prompt.</span></span>

    ```javascript
    case 13001:
        getDataWithToken({ forceAddAccount: true });
        break;
    ```

1. <span data-ttu-id="8af07-237">`TODO3` を次のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="8af07-237">Replace `TODO3` with the following code.</span></span> <span data-ttu-id="8af07-238">エラー 13002 は、ユーザーのサインインまたは同意が中断された場合に発生します。</span><span class="sxs-lookup"><span data-stu-id="8af07-238">Error 13002 occurs when user's sign-in or consent was aborted.</span></span> <span data-ttu-id="8af07-239">ユーザーに対して 1 回だけ再試行を求めます。</span><span class="sxs-lookup"><span data-stu-id="8af07-239">Ask the user to try again but no more than once again.</span></span>

    ```javascript
    case 13002:
        if (timesGetOneDriveFilesHasRun < 2) {
            showResult(['Your sign-in or consent was aborted before completion. Please try that operation again.']);
        } else {
            logError(result);
        }          
        break; 
    ```

1. <span data-ttu-id="8af07-240">`TODO4` を次のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="8af07-240">Replace `TODO4` with the following code.</span></span> <span data-ttu-id="8af07-241">エラー 13003 は、ユーザーが職場または学校アカウントと、Microsoft アカウントのどちらでもないアカウントでログインしている場合に発生します。</span><span class="sxs-lookup"><span data-stu-id="8af07-241">Error 13003 occurs when user is logged in with an account that is neither work or school, nor Microsoft account.</span></span> <span data-ttu-id="8af07-242">ユーザーに対して、サインアウトしてからサポートされているアカウントの種類でもう一度サインインするように求めます。</span><span class="sxs-lookup"><span data-stu-id="8af07-242">Ask the user to sign-out and then in again with a supported account type.</span></span>

    ```javascript
    case 13003: 
        showResult(['Please sign out of Office and sign in again with a work or school account, or Microsoft account. Other kinds of accounts, like corporate domain accounts do not work.']);
        break;   
    ```

    > [!NOTE]
    > <span data-ttu-id="8af07-243">エラー 13004 と 13005 は、開発時にのみ発生するため、このメソッドでは処理しません。</span><span class="sxs-lookup"><span data-stu-id="8af07-243">Errors 13004 and 13005 are not handled in this method because they should only occur in development.</span></span> <span data-ttu-id="8af07-244">これらは、ランタイム コードで修正できるものではなく、エンド ユーザーに報告しても意味がありません。</span><span class="sxs-lookup"><span data-stu-id="8af07-244">They cannot be fixed by runtime code and there would be no point in reporting them to an end user.</span></span>

1. <span data-ttu-id="8af07-p134">`TODO5` を次のコードと置き換えます。エラー 13006 は、Office ホストで未指定のエラーがある場合に発生します。ホストが不安定な状態にあることを示している可能性があります。ユーザーに Office の再起動を求めます。</span><span class="sxs-lookup"><span data-stu-id="8af07-p134">Replace `TODO5` with the following code. Error 13006 occurs when there has been an unspecified error in the Office host that may indicate that the host is in an unstable state. Ask the user to restart Office.</span></span>

    ```javascript
    case 13006:
        showResult(['Please save your work, sign out of Office, close all Office applications, and restart this Office application.']);
        break;        
    ```

1. <span data-ttu-id="8af07-248">`TODO6` を次のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="8af07-248">Replace `TODO6` with the following code.</span></span> <span data-ttu-id="8af07-249">エラー 13007 は、Office ホストの AAD との相互作用に問題があり、ホストがアドイン Web サービス/アプリケーションへのアクセス トークンを取得できない場合に発生します。</span><span class="sxs-lookup"><span data-stu-id="8af07-249">Error 13007 occurs when something has gone wrong with the Office host's interaction with AAD so the host cannot get an access token to the add-ins web service/application.</span></span> <span data-ttu-id="8af07-250">ネットワークに一時的な問題が発生している可能性があります。</span><span class="sxs-lookup"><span data-stu-id="8af07-250">This may be a temporary network issue.</span></span> <span data-ttu-id="8af07-251">しばらく待ってから再試行するようにユーザーに求めます。</span><span class="sxs-lookup"><span data-stu-id="8af07-251">Ask the user to try again later.</span></span>

    ```javascript
    case 13007:
        showResult(['That operation cannot be done at this time. Please try again later.']);
        break;      
    ```

1. <span data-ttu-id="8af07-p136">`TODO7` を次のコードと置き換えます。エラー 13008 は、前回の `getAccessTokenAsync` の呼び出しが完了する前に、それを呼び出す操作をユーザーがトリガーしたときに発生します。</span><span class="sxs-lookup"><span data-stu-id="8af07-p136">Replace `TODO7` with the following code. Error 13008 occurs when the user tiggered an operation that calls `getAccessTokenAsync` before a previous call of it completed.</span></span>

    ```javascript
    case 13008:
        showResult(['Please try that operation again after the current operation has finished.']);
        break;
    ```      

1. <span data-ttu-id="8af07-254">`TODO8` を次のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="8af07-254">Replace `TODO8` with the following code.</span></span> <span data-ttu-id="8af07-255">エラー 13009 は、アドインが強制的な同意をサポートしていないときに、`forceConsent` オプションを `true` に設定して `getAccessTokenAsync` を呼び出した場合に発生します。</span><span class="sxs-lookup"><span data-stu-id="8af07-255">Error 13009 occurs when the add-in does not support forcing consent, but `getAccessTokenAsync` was called with the `forceConsent` option set to `true`.</span></span> <span data-ttu-id="8af07-256">通常、この場合は、コードによって同意オプションを `false` に設定して自動的に `getAccessTokenAsync` を再実行する必要があります。</span><span class="sxs-lookup"><span data-stu-id="8af07-256">In the usual case when this happens the code should automatically re-run `getAccessTokenAsync` with the consent option set to `false`.</span></span> <span data-ttu-id="8af07-257">ただし、`forceConsent` を `true` に設定してメソッドを呼び出すこと自体が、そのオプションを `false` に設定したメソッドの呼び出しで発生したエラーに対する自動的な応答の場合もあります。</span><span class="sxs-lookup"><span data-stu-id="8af07-257">However, in some cases, calling the method with `forceConsent` set to `true` was itself an automatic response to an error in a call to the method with the option set to `false`.</span></span> <span data-ttu-id="8af07-258">その場合は、コードで再試行するのではなく、ユーザーにサインアウトしてから再度サインインするように通知する必要があります。</span><span class="sxs-lookup"><span data-stu-id="8af07-258">In that case, the code should not try again, but instead it should advise the user to sign out and sign in again.</span></span>

    ```javascript
    case 13009:
        if (triedWithoutForceConsent) {
            showResult(['Please sign out of Office and sign in again with a work or school account, or Microsoft account.']);
        } else {
            getDataWithToken({ forceConsent: false });
        }
        break;
    ```      
    
1. <span data-ttu-id="8af07-259">`TODO9` を次のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="8af07-259">Replace `TODO9` with the following code.</span></span>

    ```javascript
    default:
        logError(result);
        break;
    ```  


1. <span data-ttu-id="8af07-p138">`handleClientSideErrors` メソッドの下に、次のメソッドを追加します。このメソッドは、代理 (on-behalf-of) フローの実行時または Microsoft Graph からのデータの取得時の問題により、アドインの Web サービスで発生したエラーを処理します。</span><span class="sxs-lookup"><span data-stu-id="8af07-p138">Below the `handleClientSideErrors` method, add the following method. This method will handle errors in the add-in's web service when something goes wrong in executing the on-behalf-of flow or in getting data from Microsoft Graph.</span></span>

    ```javascript
    function handleServerSideErrors(result) {
    
        // TODO10: Parse the JSON response.

        // TODO11: Handle the case where AAD asks for an additional form of authentication.

        // TODO12: Handle missing consent and scope (permission) related issues.

        // TODO13: Handle the case where the token sent to Microsoft Graph in the request for 
        //         data is expired or invalid.

        // TODO14: Log all other server errors.
    }
    ```

1. <span data-ttu-id="8af07-262">`TODO10` を次のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="8af07-262">Replace `TODO10` with the following code.</span></span> <span data-ttu-id="8af07-263">アドインの Web サービスがアドインのクライアント側に渡すほとんどの `4xx` エラーには、その応答内に **ExceptionMessage** プロパティが含まれています。このプロパティには、AADSTS (Azure Active Directory Secure Token Service) エラー番号などのデータが格納されています。</span><span class="sxs-lookup"><span data-stu-id="8af07-263">Note that for most of the `4xx` errors that the add-in's web service will pass to the add-in's client-side, there will be an **ExceptionMessage** property in the response that contains the AADSTS (Azure Active Directory Secure Token Service) error number as well as other data.</span></span> <span data-ttu-id="8af07-264">ただし、AAD がアドインの Web サービスに追加の認証要素を求めるメッセージを送信するときには、そのメッセージに特殊な **Claims** プロパティが含まれます。このプロパティによって、どの追加要素が必要になるかが (コード番号で) 示されます。</span><span class="sxs-lookup"><span data-stu-id="8af07-264">However, when AAD sends a message to the add-in's web service asking for an additonal authentication factor, the message contains a special **Claims** property that specifies (with a code number) what additional factor is needed.</span></span> <span data-ttu-id="8af07-265">HTTP 応答を作成してクライアントに送信する ASP.NET API は、この **Claims** プロパティを認識しないため、このプロパティを応答オブジェクトに含めません。</span><span class="sxs-lookup"><span data-stu-id="8af07-265">The ASP.NET APIs that create and send HTTP Responses to clients do not know about this **Claims** property, so they do not include it in the Response object.</span></span> <span data-ttu-id="8af07-266">この後の手順で作成するサーバー側のコードでは、これに対処するために、手動で応答オブジェクトに **Claims** 値を追加しています。</span><span class="sxs-lookup"><span data-stu-id="8af07-266">Server-side code that you will create in a later step will cope with this by manually adding the **Claims** value to the Response object.</span></span> <span data-ttu-id="8af07-267">この値は、**Message** プロパティに含めるため、コードでは、そのプロパティも解析する必要があります。</span><span class="sxs-lookup"><span data-stu-id="8af07-267">This value will be in the **Message** property, so the code needs to parse out that property as well.</span></span>

    ```javascript
    var exceptionMessage = JSON.parse(result.responseText).ExceptionMessage;
    var message = JSON.parse(result.responseText).Message;
    ```

1. <span data-ttu-id="8af07-p140">`TODO11` を次のコードに置き換えます。このコードの注意点は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="8af07-p140">Replace `TODO11` with the following code. Note about this code:</span></span>

    * <span data-ttu-id="8af07-270">エラー 50076 は、Microsoft Graph が認証の追加フォームを必要とする場合に発生します。</span><span class="sxs-lookup"><span data-stu-id="8af07-270">Error 50076 occurs when Microsoft Graph requires an additional form of authentication.</span></span>
    * <span data-ttu-id="8af07-271">Office ホストは、`authChallenge` オプションとして **Claims** 値を使用して新しいトークンを取得します。</span><span class="sxs-lookup"><span data-stu-id="8af07-271">The Office host should get a new token with the **Claims** value as the `authChallenge` option.</span></span> <span data-ttu-id="8af07-272">これにより、認証の必要なフォームをすべてユーザーに表示するように AAD に指示します。</span><span class="sxs-lookup"><span data-stu-id="8af07-272">This tells AAD to prompt the user for all required forms of authentication.</span></span> 

    ```javascript
    if (message) {
        if (message.indexOf("AADSTS50076") !== -1) {
            var claims = JSON.parse(message).Claims;
            var claimsAsString = JSON.stringify(claims);
            getDataWithToken({ authChallenge: claimsAsString });
        }
    }    
    ```

1. <span data-ttu-id="8af07-273">`TODO12` を次のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="8af07-273">Replace `TODO12` with the following code.</span></span> <span data-ttu-id="8af07-274">次のいくつかの手順では、このコード内にある 3 つの `TODO` を*内部*の条件ブロックに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="8af07-274">You will replace the three `TODO`s in this code with an *inner* conditional block in the next few steps.</span></span>

    ```javascript
    else if (exceptionMessage) {

        // TODO12A: Handle the case where consent has not been granted, or has been revoked.

        // TODO12B: Handle the case where an invalid scope (permission) was used in the on-behalf-of flow.

        // TODO12C: Handle the case where the token that the add-in's client-side sends to it's 
        //          server-side is not valid because it is missing `access_as_user` scope (permission).
    }
  
    ```


1. <span data-ttu-id="8af07-275">`TODO12A` を次のコードに置き換えます </span><span class="sxs-lookup"><span data-stu-id="8af07-275">Replace `TODO12A` with the following code.</span></span> <span data-ttu-id="8af07-276">(これにより、*内部*の条件ブロックの最初の部分が作成されます)。次のコードに注意してください。</span><span class="sxs-lookup"><span data-stu-id="8af07-276">(This creates the first part of an *inner* conditional block.) Note about this code:</span></span>

    * <span data-ttu-id="8af07-277">エラー 65001 は、1 つ以上のアクセス許可について Microsoft Graph にアクセスするための同意が与えられていない (または取り消されている) ことを意味します。</span><span class="sxs-lookup"><span data-stu-id="8af07-277">Error 65001 means that consent to access Microsoft Graph was not granted (or was revoked) for one or more permissions.</span></span> 
    * <span data-ttu-id="8af07-278">アドインでは、`forceConsent` オプションを `true` に設定して新しいトークンを取得する必要があります。</span><span class="sxs-lookup"><span data-stu-id="8af07-278">The add-in should get a new token with the `forceConsent` option set to `true`.</span></span>

    ```javascript
    if (exceptionMessage.indexOf('AADSTS65001') !== -1) {
        showResult(['Please grant consent to this add-in to access your Microsoft Graph data.']);        
        /*
            THE FORCE CONSENT OPTION IS NOT AVAILABLE IN DURING PREVIEW. WHEN SSO FOR
            OFFICE ADD-INS IS RELEASED, REMOVE THE showResult LINE ABOVE AND UNCOMMENT
            THE FOLLOWING LINE.
        */
       // getDataWithToken({ forceConsent: true });
    }    
    ```

1. <span data-ttu-id="8af07-p144">`TODO12B` を次のコードに置き換えます。このコードの注意点は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="8af07-p144">Replace `TODO12B` with the following code. Note about this code:</span></span>

    * <span data-ttu-id="8af07-p145">エラー 70011 には複数の意味があります。無効なスコープ (アクセス許可) が要求されていることを意味する場合、このアドインに重要となります。コードでは番号だけでなくエラーの説明全体を確認します。</span><span class="sxs-lookup"><span data-stu-id="8af07-p145">Error 70011 has multiple meanings. The one that matters to this add-in is when it means that an invalid scope (permission) has been requested, so the code checks for the full error description, not just the number.</span></span>
    * <span data-ttu-id="8af07-283">アドインでは、エラーを報告する必要があります。</span><span class="sxs-lookup"><span data-stu-id="8af07-283">The add-in should report the error.</span></span>

    ```javascript
     else if (exceptionMessage.indexOf("AADSTS70011: The provided value for the input parameter 'scope' is not valid.") !== -1) {
        showResult(['The add-in is asking for a type of permission that is not recognized.']);
    }    
    ```

1. <span data-ttu-id="8af07-p146">`TODO12C` を次のコードに置き換えます。このコードの注意点は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="8af07-p146">Replace `TODO12C` with the following code. Note about this code:</span></span>

    * <span data-ttu-id="8af07-286">この後の手順で作成するサーバー側のコードでは、アドインのクライアントが AAD に送信して代理 (on-behalf-of) フローで使用されるアクセス トークンに `access_as_user` スコープ (アクセス許可) が含まれていない場合に、メッセージ `Missing access_as_user` を送信します。</span><span class="sxs-lookup"><span data-stu-id="8af07-286">Server-side code that you create in a later step will send the message `Missing access_as_user` if the `access_as_user` scope (permission) is not in the access token that the add-in's client sends to AAD to be used in the on-behalf-of flow.</span></span>
    * <span data-ttu-id="8af07-287">アドインでは、エラーを報告する必要があります。</span><span class="sxs-lookup"><span data-stu-id="8af07-287">The add-in should report the error.</span></span>

    ```javascript
    else if (exceptionMessage.indexOf('Missing access_as_user.') !== -1) {
        showResult(['Microsoft Office does not have permission to get Microsoft Graph data on behalf of the current user.']);
    }    
    ```

1. <span data-ttu-id="8af07-288">`TODO13` を次のコードに置き換えます </span><span class="sxs-lookup"><span data-stu-id="8af07-288">Replace `TODO13` with the following code.</span></span> <span data-ttu-id="8af07-289">(これは*外側*の条件ブロックの一部で、`else if (exceptionMessage) {` で始まる構造の右大かっこの直後にあり、同じレベルのインデントになります)。次のコードに注意してください。</span><span class="sxs-lookup"><span data-stu-id="8af07-289">(This is part of the *outer* conditional block and should be immediately after the close bracket of the structure that begins with `else if (exceptionMessage) {` and at the same level of indentation.) Note about this code:</span></span>

    * <span data-ttu-id="8af07-p148">サーバー側のコードで使用する ID ライブラリ (Microsoft Authentication Library - MSAL) では、期限切れのトークンや無効なトークンが Microsoft Graph に送信されないようにする必要があります。ただし、その事態が発生した場合は、アドインの Web サービスに Microsoft Graph から返されるエラーにコード `InvalidAuthenticationToken` が含まれています。後の手順で作成するサーバー側のコードは、このメッセージをアドインのクライアントに中継します。</span><span class="sxs-lookup"><span data-stu-id="8af07-p148">The identity library that you will be using in the server-side code (Microsoft Authentication Library - MSAL) should ensure that no expired or invalid token is sent to Microsoft Graph; but if it does happen, the error that is returned to the add-in's web service from Microsoft Graph has the code `InvalidAuthenticationToken`. Server-side code you will create in a latter step will relay this message to the add-in's client.</span></span>
    * <span data-ttu-id="8af07-292">この場合、アドインはカウンター変数とフラグ変数をリセットしてから、ボタン ハンドラー メソッドを再呼び出しすることで、認証プロセス全体を最初から開始する必要があります。</span><span class="sxs-lookup"><span data-stu-id="8af07-292">In this case, the add-in should start the entire authentication process over by resetting the counter and flag varibles, and then re-calling the button handler method.</span></span>

    ```javascript
    // If the token sent to MS Graph is expired or invalid, start the whole process over.
    else if (result.code === 'InvalidAuthenticationToken') {
        timesGetOneDriveFilesHasRun = 0;
        triedWithoutForceConsent = false;
        getOneDriveFiles();
    }    
    ```

1. <span data-ttu-id="8af07-293">`TODO14` を次のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="8af07-293">Replace `TODO14` with the following code.</span></span>

    ```javascript
    else {
        logError(result);
    }    
    ```

1. <span data-ttu-id="8af07-294">ファイルを保存して閉じます。</span><span class="sxs-lookup"><span data-stu-id="8af07-294">Save and close the file.</span></span>

## <a name="code-the-server-side"></a><span data-ttu-id="8af07-295">サーバー側のコードを作成する</span><span class="sxs-lookup"><span data-stu-id="8af07-295">Code the server side</span></span>

### <a name="configure-the-owin-middleware"></a><span data-ttu-id="8af07-296">OWIN ミドルウェアを構成する</span><span class="sxs-lookup"><span data-stu-id="8af07-296">Configure the OWIN middleware</span></span>

1. <span data-ttu-id="8af07-297">プロジェクトのルートにある Startup.cs を開きます。</span><span class="sxs-lookup"><span data-stu-id="8af07-297">Open the Startup.cs file in the root of the project.</span></span>

1. <span data-ttu-id="8af07-p149">Startup クラスの宣言にキーワード `partial` を追加します (まだ追加されていない場合)。これは、次のようになります。</span><span class="sxs-lookup"><span data-stu-id="8af07-p149">Add the keyword `partial` to the declaration of the Startup class, if it is not already there. It should look like this:</span></span>

    `public partial class Startup`

1. <span data-ttu-id="8af07-p150">`Configuration` メソッドの本文に、次に示す行を追加します。`ConfigureAuth` メソッドは、この後の手順で作成します。</span><span class="sxs-lookup"><span data-stu-id="8af07-p150">Add the following line to the body of the `Configuration` method. You create the `ConfigureAuth` method in a later step.</span></span>

    `ConfigureAuth(app);`

1. <span data-ttu-id="8af07-302">ファイルを保存して閉じます。</span><span class="sxs-lookup"><span data-stu-id="8af07-302">Save and close the file.</span></span>

1. <span data-ttu-id="8af07-303">**App_Start** フォルダーを右クリックして、**[追加] > [クラス]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="8af07-303">Right-click the **App_Start** folder and select **Add > Class**.</span></span>

1. <span data-ttu-id="8af07-304">**[新しい項目の追加]** ダイアログで、ファイルに「**Startup.Auth.cs**」という名前を付けて **[追加]** をクリックします。</span><span class="sxs-lookup"><span data-stu-id="8af07-304">In the **Add new item** dialog name the file **Startup.Auth.cs** and then click **Add**.</span></span>

1. <span data-ttu-id="8af07-305">新しいファイルで名前空間の名前を `Office_Add_in_ASPNET_SSO_WebAPI` に短縮します。</span><span class="sxs-lookup"><span data-stu-id="8af07-305">Shorten the namespace name in the new file to `Office_Add_in_ASPNET_SSO_WebAPI`.</span></span>

1. <span data-ttu-id="8af07-306">ファイルの先頭に、次に示す `using` ステートメントがすべて揃っていることを確認します。</span><span class="sxs-lookup"><span data-stu-id="8af07-306">Ensure that all of the following `using` statements are at the top of the file.</span></span>

    ```csharp
    using Owin;
    using System.IdentityModel.Tokens;
    using System.Configuration;
    using Microsoft.Owin.Security.OAuth;
    using Microsoft.Owin.Security.Jwt;
    using Office_Add_in_ASPNET_SSO_WebAPI.App_Start;
    ```

1. <span data-ttu-id="8af07-p151">`Startup` クラスの宣言にキーワード `partial` を追加します (まだ追加されていない場合)。これは、次のようになります。</span><span class="sxs-lookup"><span data-stu-id="8af07-p151">Add the keyword `partial` to the declaration of the `Startup` class, if it is not already there. It should look like this:</span></span>

    `public partial class Startup`

1. <span data-ttu-id="8af07-p152">次に示すメソッドを `Startup` クラスに追加します。このメソッドでは、クライアント側の Home.js ファイルの `getData` メソッドから渡されたアクセス トークンを OWIN ミドルウェアで検証する方法を指定します。承認プロセスは、`[Authorize]` 属性で修飾された Web API エンドポイントが呼び出されたときには必ずトリガーされます。</span><span class="sxs-lookup"><span data-stu-id="8af07-p152">Add the following method to the `Startup` class. This method specifies how the OWIN middleware will validate the access tokens that are passed to it from the `getData` method in the client-side Home.js file. The authorization process is triggered whenever a Web API endpoint that is decorated with the `[Authorize]` attribute is called.</span></span>

    ```csharp
    public void ConfigureAuth(IAppBuilder app)
    {
        // TODO3: Configure the validation settings
        // TODO4: Specify the type of authorization and the discovery endpoint
        // of the secure token service.
    }
    ```

1. <span data-ttu-id="8af07-312">TODO3 を次のように置き換えます。</span><span class="sxs-lookup"><span data-stu-id="8af07-312">Replace the TODO3 with the following.</span></span> <span data-ttu-id="8af07-313">このコードの注意点は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="8af07-313">Note about this code:</span></span>

    * <span data-ttu-id="8af07-314">このコードでは、Office ホストから得られるアクセス トークン (`getData` のクライアント側呼び出しによって渡されるトークン) で指定された対象ユーザーとトークン発行者が web.config で指定された値と一致する必要があることを OWIN に指示します。</span><span class="sxs-lookup"><span data-stu-id="8af07-314">The code instructs OWIN to ensure that the audience and token issuer specified in the access token that comes from the Office host (and is passed on by the client-side call of `getData`) must match the values specified in the web.config.</span></span>
    * <span data-ttu-id="8af07-p154">`SaveSigninToken` を `true` に設定することで、OWIN は Office ホストからの Raw トークンを保存するようになります。これは、アドインが「代理」フローで Microsoft Graph へのアクセス トークンを取得するために必要になります。</span><span class="sxs-lookup"><span data-stu-id="8af07-p154">Setting `SaveSigninToken` to `true` causes OWIN to save the raw token from the Office host. The add-in needs it to obtain an access token to Microsoft Graph with the “on behalf of” flow.</span></span>
    * <span data-ttu-id="8af07-p155">OWIN ミドルウェアでは、スコープは検証されません。`access_as_user` が含まれている必要があるアクセス トークンのスコープは、コントローラーで検証されます。</span><span class="sxs-lookup"><span data-stu-id="8af07-p155">Scopes are not validated by the OWIN middleware. The scopes of the access token, which should include `access_as_user`, is validated in the controller.</span></span>

    ```csharp
    var tvps = new TokenValidationParameters
        {
            ValidAudience = ConfigurationManager.AppSettings["ida:Audience"],
            ValidIssuer = ConfigurationManager.AppSettings["ida:Issuer"],
            SaveSigninToken = true
        };
    ```

1. <span data-ttu-id="8af07-p156">TODO4 を次のように置き換えます。このコードの注意点は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="8af07-p156">Replace TODO4 with the following. Note about this code:</span></span>

    * <span data-ttu-id="8af07-321">より一般的な `UseWindowsAzureActiveDirectoryBearerAuthentication` は Azure AD V2 エンドポイントに準拠していないため、その代わりとしてメソッド `UseOAuthBearerAuthentication` が呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="8af07-321">The method `UseOAuthBearerAuthentication` is called instead of the more common `UseWindowsAzureActiveDirectoryBearerAuthentication` because the latter is not compatible with the Azure AD V2 endpoint.</span></span>
    * <span data-ttu-id="8af07-322">このメソッドに渡される探索 URL は、Office ホストから受け取ったアクセス トークンの署名の検証に必要になるキーを取得するための方法を OWIN ミドルウェアが取得する場所になります。</span><span class="sxs-lookup"><span data-stu-id="8af07-322">The discovery URL that is passed to the method is where the OWIN middleware obtains instructions for getting the key it needs to verify the signature on the access token received from the Office host.</span></span>

    ```csharp
    app.UseOAuthBearerAuthentication(new OAuthBearerAuthenticationOptions
        {
            AccessTokenFormat = new JwtFormat(tvps, new OpenIdConnectCachingSecurityTokenProvider("https://login.microsoftonline.com/common/v2.0/.well-known/openid-configuration"))
        });
    ```

1. <span data-ttu-id="8af07-323">ファイルを保存して閉じます。</span><span class="sxs-lookup"><span data-stu-id="8af07-323">Save and close the file.</span></span>

### <a name="create-the-apivalues-controller"></a><span data-ttu-id="8af07-324">/api/values コントローラーを作成する</span><span class="sxs-lookup"><span data-stu-id="8af07-324">Create the /api/values controller</span></span>

1. <span data-ttu-id="8af07-325">ファイル **Controllers\ValueController.cs** を開きます。</span><span class="sxs-lookup"><span data-stu-id="8af07-325">Open the file **Controllers\ValueController.cs**.</span></span>

2. <span data-ttu-id="8af07-326">ファイルの先頭に、次に示す `using` ステートメントがあることを確認します。</span><span class="sxs-lookup"><span data-stu-id="8af07-326">Ensure that the following `using` statements are at the top of the file.</span></span>

    ```csharp
    using Microsoft.Identity.Client;
    using System.IdentityModel.Tokens;
    using System.Collections.Generic;
    using System.Configuration;
    using System.Linq;
    using System.Security.Claims;
    using System.Threading.Tasks;
    using System.Web.Http;
    using System;
    using System.Net;
    using System.Net.Http;
    using Office_Add_in_ASPNET_SSO_WebAPI.Helpers;
    using Office_Add_in_ASPNET_SSO_WebAPI.Models;
    ```

3. <span data-ttu-id="8af07-p157">`ValuesController` を宣言している行のすぐ上に、属性 `[Authorize]` を追加します。これにより、アドインはコントローラー メソッドが呼び出されたときに、最後の手順で構成した承認プロセスを必ず実行するようになります。アドインへの有効なアクセス トークンを持つ呼び出し元のみが、コントローラーのメソッドを起動できます。</span><span class="sxs-lookup"><span data-stu-id="8af07-p157">Just above the line that declares the `ValuesController`, add the `[Authorize]` attribute. This ensures that your add-in will run the authorization process that you configured in the last procedure whenever a controller method is called. Only callers with a valid access token to your add-in can invoke the methods of the controller.</span></span>

    > [!NOTE]
    > <span data-ttu-id="8af07-330">運用環境の ASP.NET MVC Web API サービスには、1 つ以上のカスタム **FilterAttribute** クラスに代理 (on-behalf-of) フロー用のカスタム ロジックを用意する必要があります。</span><span class="sxs-lookup"><span data-stu-id="8af07-330">A production ASP.NET MVC Web API service should have custom logic for the on-behalf-of flow in one or more custom **FilterAttribute** classes.</span></span> <span data-ttu-id="8af07-331">この学習用サンプルでは、メイン コントローラーにロジックを配置して、認証とデータのフェッチ ロジックの全体的なフローを簡単に把握できるようにしています。</span><span class="sxs-lookup"><span data-stu-id="8af07-331">This educational sample puts the logic in the main controller so that the entire flow of the authorization and data fetching logic can be easily followed.</span></span> <span data-ttu-id="8af07-332">さらに、このサンプルが「[Azure Samples](https://github.com/Azure-Samples/)」の承認サンプルのパターンと一致するようになります。</span><span class="sxs-lookup"><span data-stu-id="8af07-332">This also makes the sample consistent with the pattern of authorization samples in [Azure Samples](https://github.com/Azure-Samples/).</span></span>    

4. <span data-ttu-id="8af07-333">次のメソッドを `ValuesController` に追加します。</span><span class="sxs-lookup"><span data-stu-id="8af07-333">Add the following method to the `ValuesController`.</span></span> <span data-ttu-id="8af07-334">戻り値は、`Task<IEnumerable<string>>` ではなく `GET api/values` メソッドでより一般的な `Task<HttpResponseMessage>` になる点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="8af07-334">Note that the return value is `Task<HttpResponseMessage>` instead of `Task<IEnumerable<string>>` as would be more common for a `GET api/values` method.</span></span> <span data-ttu-id="8af07-335">これは、カスタムの承認ロジックがコントローラー内にあることの副作用です。そのロジックの一部のエラー条件では、HTTP 応答オブジェクトをアドインのクライアントに送信することが必要になります。</span><span class="sxs-lookup"><span data-stu-id="8af07-335">This is a side effect of that fact that our custom authorization logic will be in the controller: some error conditions in that logic require that an HTTP Response object be sent to the add-in's client.</span></span> 

    ```csharp
    // GET api/values
    public async Task<HttpResponseMessage> Get()
    {
        // TODO1: Validate the scopes of the access token.
    }
    ```

5. <span data-ttu-id="8af07-336">`TODO1` を次のコードに置き換えて、`access_as_user` を含むトークンで指定されているスコープを検証します。</span><span class="sxs-lookup"><span data-stu-id="8af07-336">Replace `TODO1` with the following code to validate that the scopes that are specified in the token include `access_as_user`.</span></span>

    ```csharp
    string[] addinScopes = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/scope").Value.Split(' ');
    if (addinScopes.Contains("access_as_user"))
    {
        // TODO2: Assemble all the information that is needed to get a token for Microsoft Graph using the "on behalf of" flow.
        // TODO3: Get the access token for Microsoft Graph.
        // TODO4: Get the names of files and folders in OneDrive by using the Microsoft Graph API.
        // TODO5: Remove excess information from the data and send the data to the client.
    }
    return SendErrorToClient(HttpStatusCode.Unauthorized, null, "Missing access_as_user.");
    ```

    > [!NOTE]
    > <span data-ttu-id="8af07-p160">`access_as_user` スコープのみを使用して、Office アドインの代理 (on-behalf-of) フローを処理する API を承認する必要があります。サービス内の他の API は、独自のスコープ要件が必要です。これにより、Office が取得するトークンでアクセスできるものが制限されます。</span><span class="sxs-lookup"><span data-stu-id="8af07-p160">You should only use the `access_as_user` scope to authorize the API that handles the on-behalf-of flow for Office Add-ins. Other APIs in your service should have their own scope requirements. This limits what can be accessed with the tokens that Office acquires.</span></span>

6. <span data-ttu-id="8af07-p161">`TODO2` を次のコードに置き換えます。このコードの注意点は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="8af07-p161">Replace `TODO2` with the following code. Note about this code:</span></span>
    * <span data-ttu-id="8af07-341">このコードでは、Office ホストから受け取った Raw アクセス トークンを別のメソッドに渡される `UserAssertion` オブジェクトに変換します。</span><span class="sxs-lookup"><span data-stu-id="8af07-341">It turns the raw access token received from the Office host into a `UserAssertion` object that will be passed to another method.</span></span>
    * <span data-ttu-id="8af07-p162">アドインは、Office ホストとユーザーがアクセスする必要のあるリソース (または対象ユーザー) の役割を果たさなくなります。この時点で、それ自体が Microsoft Graph にアクセスする必要があるクライアントになります。`ConfidentialClientApplication` は MSAL の「クライアント コンテキスト」オブジェクトになります。</span><span class="sxs-lookup"><span data-stu-id="8af07-p162">Your add-in is no longer playing the role of a resource (or audience) to which the Office host and user need access. Now it is itself a client that needs access to Microsoft Graph. `ConfidentialClientApplication` is the MSAL “client context” object.</span></span>
    * <span data-ttu-id="8af07-p163">`ConfidentialClientApplication` コンストラクターへの 3 番目のパラメーターはリダイレクト URL です。これは、実際には「代理」フローで使用されることはありませんが、正しい URL を使用することをお勧めします。4 番目と 5 番目のパラメーターは、永続ストアを定義するために使用できます。このストアにより、有効期限が切れていないトークンをアドインの異なるセッション間で再使用できるようになります。このサンプルでは、永続ストアは実装していません。</span><span class="sxs-lookup"><span data-stu-id="8af07-p163">The third parameter to the `ConfidentialClientApplication` constructor is a redirect URL which is not actually used in the “on behalf of” flow, but it is a good practice to use the correct URL. The fourth and fifth parameters can be used to define a persistent store that would enable the reuse of unexpired tokens across different sessions with the add-in. This sample does not implement any persistent storage.</span></span>
    * <span data-ttu-id="8af07-348">MSAL では `openid`、`offline_access` の各スコープが機能することが必要ですが、コードがこれらを重複して要求するとエラーがスローされます。</span><span class="sxs-lookup"><span data-stu-id="8af07-348">MSAL requires the `openid` and `offline_access` scopes to function, but it throws an error if your code redundantly requests them.</span></span> <span data-ttu-id="8af07-349">コードが `profile` を要求した場合にもエラーがスローされます。それは、実際には Office ホスト アプリケーションがアドインの Web アプリケーションに対しトークンを取得するときだけに使用します。</span><span class="sxs-lookup"><span data-stu-id="8af07-349">It will also throw an error if your code requests `profile`, which is really only used when the Office host application gets the token to your add-in's web application.</span></span> <span data-ttu-id="8af07-350">そのため、`Files.Read.All` のみが明示的に要求されます。</span><span class="sxs-lookup"><span data-stu-id="8af07-350">So only `Files.Read.All` is explicitly requested.</span></span>

    ```csharp
    var bootstrapContext = ClaimsPrincipal.Current.Identities.First().BootstrapContext as BootstrapContext;
    UserAssertion userAssertion = new UserAssertion(bootstrapContext.Token);
    ClientCredential clientCred = new ClientCredential(ConfigurationManager.AppSettings["ida:Password"]);
    ConfidentialClientApplication cca =
                    new ConfidentialClientApplication(ConfigurationManager.AppSettings["ida:ClientID"],
                                                      "https://localhost:44355", clientCred, null, null);
    string[] graphScopes = { "Files.Read.All" };
    ```

7. <span data-ttu-id="8af07-p165">`TODO3` を次のコードに置き換えます。このコードの注意点は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="8af07-p165">Replace `TODO3` with the following code. Note about this code:</span></span>

    * <span data-ttu-id="8af07-p166">`ConfidentialClientApplication.AcquireTokenOnBehalfOfAsync` メソッドは、最初にメモリ内の MSAL キャッシュで一致するアクセス トークンを探します。それが見つからなかった場合にのみ、Azure AD V2 エンドポイントとの「代理」フローを開始します。</span><span class="sxs-lookup"><span data-stu-id="8af07-p166">The `ConfidentialClientApplication.AcquireTokenOnBehalfOfAsync` method will first look in the MSAL cache, which is in memory, for a matching access token. Only if there isn't one, does it initiate the "on behalf of" flow with the Azure AD V2 endpoint.</span></span>
    * <span data-ttu-id="8af07-355">MS Graph リソースが多要素認証を必要とし、ユーザーがまだそれを提供していない場合、AAD は Claims プロパティが含まれている例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="8af07-355">If multi-factor authentication is required by the MS Graph resource and the user has not yet provided it, AAD will throw an exception containing a Claims property.</span></span>
    * <span data-ttu-id="8af07-p167">Claims プロパティの値は、クライアントに渡す必要があります。クライアントは、その値を Office ホストに渡します。Office ホストは、その値を新しいトークンの要求に含めます。AAD は、認証のすべての必要なフォームをユーザーに示します。</span><span class="sxs-lookup"><span data-stu-id="8af07-p167">The Claims property value must be passed to the client which will pass it to the Office host, which will then include it in a request for a new token. AAD will prompt the user for all required forms of authentication.</span></span>
    * <span data-ttu-id="8af07-358">`MsalServiceException` 以外の種類の例外は、意図的にキャッチしていないため、`500 Server Error` メッセージとしてクライアントに伝達されます。</span><span class="sxs-lookup"><span data-stu-id="8af07-358">Any exceptions that are not of type `MsalServiceException` are intentionally not caught, so they will propagate to the client as `500 Server Error` messages.</span></span>

    ```csharp
    AuthenticationResult result = null;
    try
    {
        result = await cca.AcquireTokenOnBehalfOfAsync(graphScopes, userAssertion, "https://login.microsoftonline.com/common/oauth2/v2.0");
    }
    catch (MsalServiceException e)
    {        
        // TODO3a: Handle request for multi-factor authentication.
        // TODO3b: Handle lack of consent.
        // TODO3c: Handle invalid scope (permission).
        // TODO3d: Handle all other MsalServiceExceptions.
    }
    ```

8. <span data-ttu-id="8af07-p168">`TODO3a` を次のコードに置き換えます。このコードの注意点は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="8af07-p168">Replace `TODO3a` with the following code. Note about this code:</span></span>

    * <span data-ttu-id="8af07-p169">MS Graph リソースが多要素認証を必要としているときに、その認証をユーザーがまだ指定していない場合、AAD はエラー AADSTS50076 と **Claims** プロパティを含む "400 Bad Request" を返します。MSAL は **MsalUiRequiredException** (**MsalServiceException** から継承) をこの情報とともにスローします。</span><span class="sxs-lookup"><span data-stu-id="8af07-p169">If multi-factor authentication is required by the MS Graph resource and the user has not yet provided it, AAD will return "400 Bad Request" with error AADSTS50076 and a **Claims** property. MSAL throws a **MsalUiRequiredException** (which inherits from **MsalServiceException**) with this information.</span></span> 
    * <span data-ttu-id="8af07-p170">**Claims** プロパティの値は、クライアントに渡す必要があります。クライアントは、その値を Office ホストに渡します。Office ホストは、その値を新しいトークンの要求に含めます。AAD は、認証のすべての必要なフォームのための指示をユーザーに示します。</span><span class="sxs-lookup"><span data-stu-id="8af07-p170">The **Claims** property value must be passed to the client which should pass it to the Office host, which then includes it in a request for a new token. AAD will prompt the user for all required forms of authentication.</span></span>
    * <span data-ttu-id="8af07-365">例外から HTTP 応答を作成する API は、**Claims** プロパティを認識しないため、このプロパティを応答オブジェクトに含めません。</span><span class="sxs-lookup"><span data-stu-id="8af07-365">The APIs that create HTTP Responses from exceptions don't know about the **Claims** property, so they don't include it in the response object.</span></span> <span data-ttu-id="8af07-366">これが含まれたメッセージを手動で作成する必要があります。</span><span class="sxs-lookup"><span data-stu-id="8af07-366">We have to manually create a message that includes it.</span></span> <span data-ttu-id="8af07-367">ただし、カスタムの **Message** プロパティは **ExceptionMessage** プロパティの作成を妨げるため、クライアントがエラー ID `AADSTS50076` を取得するには、その ID をカスタムの **Message** に追加する以外に方法はありません。</span><span class="sxs-lookup"><span data-stu-id="8af07-367">A custom **Message** property, however, blocks the creation of an **ExceptionMessage** property, so the only way to get the error ID `AADSTS50076` to the client is to add it to the custom **Message**.</span></span> <span data-ttu-id="8af07-368">クライアントの JavaScript では、応答に **Message** または **ExceptionMessage** が含まれているかどうかを検出する必要があるため、どちらを読み取るかを認識します。</span><span class="sxs-lookup"><span data-stu-id="8af07-368">JavaScript in the client will need to discover if a response has a **Message** or **ExceptionMessage**, so it knows which to read.</span></span>
    * <span data-ttu-id="8af07-369">カスタム メッセージは、JSON として書式設定されているため、クライアント側の JavaScript は既知の `JSON` オブジェクトのメソッドでメッセージを解析できます。</span><span class="sxs-lookup"><span data-stu-id="8af07-369">The custom message is formatted as JSON so that the client-side JavaScript can parse it with well-known `JSON` object methods.</span></span>
    * <span data-ttu-id="8af07-370">`SendErrorToClient` メソッドは、この後の手順で作成します。</span><span class="sxs-lookup"><span data-stu-id="8af07-370">You will create the `SendErrorToClient` method in a later step.</span></span> <span data-ttu-id="8af07-371">2 番目のパラメーターは、**Exception** オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="8af07-371">It's second parameter is an **Exception** object.</span></span> <span data-ttu-id="8af07-372">この場合、コードは `null` を渡します。これは、**Exception** オブジェクトが含まれていることで、生成される HTTP 応答には **Message** プロパティが含められなくなるためです。</span><span class="sxs-lookup"><span data-stu-id="8af07-372">In this case, the code passes `null` because including the **Exception** object blocks the inclusion of the **Message** property in the HTTP Response that is generated.</span></span>

    ```csharp
    if (e.Message.StartsWith("AADSTS50076")) {
        string responseMessage = String.Format("{{\"AADError\":\"AADSTS50076\",\"Claims\":{0}}}", e.Claims);
        return SendErrorToClient(HttpStatusCode.Forbidden, null, responseMessage);
    }
    ```

9. <span data-ttu-id="8af07-p173">`TODO3b` と `TODO3c` を次のコードに置き換えます。このコードの注意点は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="8af07-p173">Replace `TODO3b` and `TODO3c` with the following code. Note about this code:</span></span>

    * <span data-ttu-id="8af07-375">AAD の呼び出しにユーザーまたはテナント管理者のどちらも同意していない (または同意が取り消された) スコープ (アクセス許可) が少なくとも 1 つ含まれていると、</span><span class="sxs-lookup"><span data-stu-id="8af07-375">If the call to AAD contained at least one scope (permission) for which neither the user nor a tenant administrator has consented (or consent was revoked).</span></span> <span data-ttu-id="8af07-376">AAD はエラー `AADSTS65001` と共に "400 Bad Request" を返します。</span><span class="sxs-lookup"><span data-stu-id="8af07-376">AAD will return "400 Bad Request" with error `AADSTS65001`.</span></span> <span data-ttu-id="8af07-377">MSAL は、この情報と共に **MsalUiRequiredException** をスローします。</span><span class="sxs-lookup"><span data-stu-id="8af07-377">MSAL throws a **MsalUiRequiredException** with this information.</span></span> <span data-ttu-id="8af07-378">クライアントは、オプション `{ forceConsent: true }` を使用して `getAccessTokenAsync` を再呼び出しする必要があります。</span><span class="sxs-lookup"><span data-stu-id="8af07-378">The client should re-call `getAccessTokenAsync` with the option `{ forceConsent: true }`.</span></span>
    *  <span data-ttu-id="8af07-379">AAD の呼び出しに AAD が認識しないスコープが少なくとも 1 つ含まれていると、AAD はエラー `AADSTS70011` と共に "400 Bad Request" を返します。</span><span class="sxs-lookup"><span data-stu-id="8af07-379">If the call to AAD contained at least one scope that AAD does not recognize, AAD returns "400 Bad Request" with error `AADSTS70011`.</span></span> <span data-ttu-id="8af07-380">MSAL は、この情報と共に **MsalUiRequiredException** をスローします。</span><span class="sxs-lookup"><span data-stu-id="8af07-380">MSAL throws a **MsalUiRequiredException** with this information.</span></span> <span data-ttu-id="8af07-381">クライアントは、ユーザーに通知する必要があります。</span><span class="sxs-lookup"><span data-stu-id="8af07-381">The client should inform the user.</span></span>
    *  <span data-ttu-id="8af07-382">すべての説明が含まれている理由は、別の条件で 70011 が返されたときに、このアドインでは無効なスコープの存在を意味する場合のみを処理する必要があるためです。</span><span class="sxs-lookup"><span data-stu-id="8af07-382">The entire description is included beause 70011 is returned in other conditions and we it should only be handled in this add-in when it means that there is an invalid scope.</span></span> 
    *  <span data-ttu-id="8af07-p176">**MsalUiRequiredException** オブジェクトが `SendErrorToClient` に渡されます。これにより、エラー情報を格納している **ExceptionMessage** プロパティが HTTP 応答に含まれるようにします。</span><span class="sxs-lookup"><span data-stu-id="8af07-p176">The **MsalUiRequiredException** object is passed to `SendErrorToClient`. This ensures that an **ExceptionMessage** property that contains the error information is included in the HTTP Response.</span></span>
    *  <span data-ttu-id="8af07-385">カスタム メッセージは存在しないため、3 番目のパラメーターでは `null` が渡されます。</span><span class="sxs-lookup"><span data-stu-id="8af07-385">There is no custom message, so `null` is passed for the third parameter.</span></span>

    ```csharp
    if ((e.Message.StartsWith("AADSTS65001"))
    || (e.Message.StartsWith("AADSTS70011: The provided value for the input parameter 'scope' is not valid.")))
    {
        return SendErrorToClient(HttpStatusCode.Forbidden, e, null);
    }
    ```

10. <span data-ttu-id="8af07-386">`TODO3d` を次のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="8af07-386">Replace `TODO3d` with the following code.</span></span> <span data-ttu-id="8af07-387">このコードでは、**HttpStatusCode.Forbidden** (401) によるカスタムの HTTP 応答で例外を中継するのではなく、例外を再スローしています。</span><span class="sxs-lookup"><span data-stu-id="8af07-387">Note that the code rethrows the exception instead of relaying it in a custom HTTP Response with **HttpStatusCode.Forbidden** (401).</span></span> <span data-ttu-id="8af07-388">これにより、ASP.NET はステータス "500 Server Error" による独自の HTTP 応答を送信するようになります。</span><span class="sxs-lookup"><span data-stu-id="8af07-388">The effect of this is that the ASP.NET will send its own HTTP Response with status "500 Server Error".</span></span>

    ```csharp
    else
    {
        throw e;
    }  
    ```

11. <span data-ttu-id="8af07-p178">`TODO4` を次のように置き換えます。このコードの注意点は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="8af07-p178">Replace `TODO4` with the following. Note about this code:</span></span>

    * <span data-ttu-id="8af07-p179">`GraphApiHelper` クラスと `ODataHelper` クラスは、**[Helpers]** フォルダー内のファイルで定義されています。`OneDriveItem` クラスは、**[Models]** フォルダー内のファイルで定義されています。これらのクラスについての詳しい説明は、承認や SSO に関連していないため、この記事の対象外になります。</span><span class="sxs-lookup"><span data-stu-id="8af07-p179">The `GraphApiHelper` and `ODataHelper` classes are defined in files in the **Helpers** folder. The `OneDriveItem` class is defined in a file in the **Models** folder. Detailed discussion of these classes is not relevant to authorization or SSO, so it is out-of-scope for this article.</span></span>
    * <span data-ttu-id="8af07-394">実際に必要なデータのみを Microsoft Graph に要求することでパフォーマンスが向上します。そのため、このコードでは、` $select` クエリ パラメーターで name プロパティのみが必要なことを指定し、`$top` パラメーターで最初の 3 つのフォルダー名またはファイル名のみが必要なことを指定しています。</span><span class="sxs-lookup"><span data-stu-id="8af07-394">Performance is improved by asking Microsoft Graph for only the data actually needed, so the code uses a ` $select` query parameter to specify that we only want the name property, and a `$top` parameter to specify that we want only the first three folder or file names.</span></span>
    * <span data-ttu-id="8af07-395">Microsoft Graph に送信したトークンが無効な場合、Microsoft Graph は、コード "InvalidAuthenticationToken" を含む "401 Unauthorized" エラーを送信します。</span><span class="sxs-lookup"><span data-stu-id="8af07-395">If the token sent to Microsoft Graph is invalid, Microsoft Graph sends a "401 Unauthorized" error with the code "InvalidAuthenticationToken".</span></span> <span data-ttu-id="8af07-396">その後で、ASP.NET は **RuntimeBinderException** をスローします。</span><span class="sxs-lookup"><span data-stu-id="8af07-396">ASP.NET then throws a **RuntimeBinderException**.</span></span> <span data-ttu-id="8af07-397">これは、トークンの有効期限が切れているときにも発生しますが、MSAL では、そのような事態にならないようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="8af07-397">This is also what happens when the token is expired, although MSAL should prevent that from ever happening.</span></span> 

    ```csharp
    var fullOneDriveItemsUrl = GraphApiHelper.GetOneDriveItemNamesUrl("?$select=name&$top=3");
    IEnumerable<OneDriveItem> filesResult;
    try
    {
        filesResult = await ODataHelper.GetItems<OneDriveItem>(fullOneDriveItemsUrl, result.AccessToken);
    }
    catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException e)
    {
        return SendErrorToClient(HttpStatusCode.Unauthorized, e, null);                    
    }
    ```

12. <span data-ttu-id="8af07-p181">`TODO5` を次のように置き換えます。このコードの注意点は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="8af07-p181">Replace `TODO5` with the following. Note about this code:</span></span> 

    * <span data-ttu-id="8af07-p182">上記のコードでは OneDrive アイテムの *name* プロパティのみを要求していますが、Microsoft Graph は常に OneDrive アイテムの *eTag* プロパティを含めます。クライアントに送信するペイロードを縮小するために、次に示すコードではアイテム名のみで結果を再構築しています。</span><span class="sxs-lookup"><span data-stu-id="8af07-p182">Although the code above asked for only the *name* property of the OneDrive items, Microsoft Graph always includes the *eTag* property for OneDrive items. To reduce the payload sent to the client, the code below reconstructs the results with only the item names.</span></span>
    * <span data-ttu-id="8af07-402">3 つの OneDrive ファイルとフォルダーのリストは、"200 OK" HTTP 応答としてクライアントに送信されます。</span><span class="sxs-lookup"><span data-stu-id="8af07-402">The list of three OneDrive files and folders is sent to the client as a "200 OK" HTTP Response.</span></span>

    ```csharp
    List<string> itemNames = new List<string>();
    foreach (OneDriveItem item in filesResult)
    {
        itemNames.Add(item.Name);
    }

    var requestMessage = new HttpRequestMessage();
    requestMessage.SetConfiguration(new HttpConfiguration());
    var response = requestMessage.CreateResponse<List<string>>(HttpStatusCode.OK, itemNames); 
    return response;
    ```

13. <span data-ttu-id="8af07-403">Get メソッドの下に、次のメソッドを追加します。</span><span class="sxs-lookup"><span data-stu-id="8af07-403">Below the Get method, add the following method.</span></span> <span data-ttu-id="8af07-404">このコードの注意点は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="8af07-404">About this code note:</span></span>  

    * <span data-ttu-id="8af07-405">このメソッドは、サーバー側の例外に関する情報をクライアントに中継します。</span><span class="sxs-lookup"><span data-stu-id="8af07-405">The method relays to the client information about a server-side exception.</span></span> 
    * <span data-ttu-id="8af07-406">このメソッドに元の例外が渡されると、HttpError コンストラクターは例外オブジェクトからの情報を **ExceptionMessage** プロパティに含めます。</span><span class="sxs-lookup"><span data-stu-id="8af07-406">If the original exception is passed to the method, then the HttpError constuctor will include information from the exception object in an **ExceptionMessage** property.</span></span>  
    * <span data-ttu-id="8af07-407">例外として `null` が渡されると、HttpError コンストラクターはメッセージ パラメーターを **Message** プロパティに含めます。**ExceptionMessage** プロパティは存在しなくなります。</span><span class="sxs-lookup"><span data-stu-id="8af07-407">If `null` is passed for the exception, then the HttpError constuctor will include the message parameter in a **Message** property and there is no **ExceptionMessage** property.</span></span>

    ```csharp
    private HttpResponseMessage SendErrorToClient(HttpStatusCode statusCode, Exception e, string message)
    {
        HttpError error;
        if (e != null)
        {
            error = new HttpError(e, true);
        }
        else
        {
            error = new HttpError(message);
        }
        var requestMessage = new HttpRequestMessage();
        var errorMessage = requestMessage.CreateErrorResponse(statusCode, error);
        return errorMessage;
    }        
    ```

## <a name="run-the-add-in"></a><span data-ttu-id="8af07-408">アドインを実行する</span><span class="sxs-lookup"><span data-stu-id="8af07-408">Run the add-in</span></span>

1. <span data-ttu-id="8af07-409">結果を確認できるように、OneDrive 内にファイルがいくつかあることを確認します。</span><span class="sxs-lookup"><span data-stu-id="8af07-409">Ensure that you have some files in your OneDrive so that you can verify the results.</span></span>

1. <span data-ttu-id="8af07-p184">Visual Studio で、F5 キーを押します。PowerPoint が開き、**[ホーム]** リボンに **[SSO ASP.NET]** グループが表示されます。</span><span class="sxs-lookup"><span data-stu-id="8af07-p184">In Visual Studio, press F5. PowerPoint opens and there is an **SSO ASP.NET** group on the **Home** ribbon.</span></span>

1. <span data-ttu-id="8af07-412">このグループ内の **[アドインの表示]** ボタンをクリックすると、作業ウィンドウにアドインの UI が表示されます。</span><span class="sxs-lookup"><span data-stu-id="8af07-412">Press the **Show Add-in** button in this group to see the add-in’s UI in the task pane.</span></span>

1. <span data-ttu-id="8af07-p185">**[OneDrive からファイルを取得]** ボタンをクリックします。Office にサインインしていない場合は、サインインを求めるダイアログが表示されます。</span><span class="sxs-lookup"><span data-stu-id="8af07-p185">Press the button **Get My Files from OneDrive**. If you are not signed into Office, you'll be prompted to sign in.</span></span>
    
    > [!NOTE]
    > <span data-ttu-id="8af07-415">以前に別の ID で Office にサインオンしていて、そのときに開いたいくつかの Office アプリケーションが引き続き開いている場合、Office がその ID を確実に変更するとは限りません (PowerPoint で ID が変更済みのように表示されている場合でも)。</span><span class="sxs-lookup"><span data-stu-id="8af07-415">If you were previously signed on to Office with a different ID, and some Office applications that were open at the time are still open, Office may not reliably change your ID even if it appears to have done so in PowerPoint.</span></span> <span data-ttu-id="8af07-416">このような場合は、Microsoft Graph への呼び出しが失敗するか、以前の ID からのデータが返される可能性があります。</span><span class="sxs-lookup"><span data-stu-id="8af07-416">If this happens, the call to Microsoft Graph may fail or data from the previous ID may be returned.</span></span> <span data-ttu-id="8af07-417">これを防止するには、必ず*他のすべての Office アプリケーションを閉じて*から、**[OneDrive からファイルを取得]** を押します。</span><span class="sxs-lookup"><span data-stu-id="8af07-417">To prevent this, be sure to *close all other Office applications* before you press **Get My Files from OneDrive**.</span></span>

1. <span data-ttu-id="8af07-p187">サインインすると、ボタンの下に OneDrive のファイルとフォルダーのリストが表示されます。これには、15 秒以上かかることがあります (特に初回実行時)。</span><span class="sxs-lookup"><span data-stu-id="8af07-p187">After you are signed in, a list of your files and folders on OneDrive will appear below the button. This may take over 15 seconds, especially the first time.</span></span>
