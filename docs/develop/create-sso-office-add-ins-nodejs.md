---
title: シングル サインオンを使用する Node.js Office アドインを作成する
description: ''
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: e304813422dea5917202ed8933c9e53df18ba9de
ms.sourcegitcommit: c5daedf017c6dd5ab0c13607589208c3f3627354
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/20/2019
ms.locfileid: "30691217"
---
# <a name="create-a-nodejs-office-add-in-that-uses-single-sign-on-preview"></a><span data-ttu-id="0cc95-102">シングル サインオンを使用する Node.js Office アドインを作成する (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="0cc95-102">Create a Node.js Office Add-in that uses single sign-on (preview)</span></span>

<span data-ttu-id="0cc95-p101">ユーザーは、このサインイン プロセスを利用してユーザーを承認する Office および Office Web アドインにサインインできます。こうして承認されたユーザーは、アドインと Microsoft Graph への 2 度目のサインオンの必要がなくなります。概要については、「[Office アドインで SSO を有効化する](sso-in-office-add-ins.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="0cc95-p101">Users can sign in to Office, and your Office Web Add-in can take advantage of this sign-in process to authorize users to your add-in and to Microsoft Graph without requiring users to sign in a second time. For an overview, see [Enable SSO in an Office Add-in](sso-in-office-add-ins.md).</span></span>

<span data-ttu-id="0cc95-105">この記事では、Node.js と Express を使用して作成したアドインで、シングル サインオン (SSO) を有効化するプロセスについて手順を追って説明します。</span><span class="sxs-lookup"><span data-stu-id="0cc95-105">This article walks you through the process of enabling single sign-on (SSO) in an add-in that is built with Node.js and Express.</span></span>

> [!NOTE]
> <span data-ttu-id="0cc95-106">ASP.NET ベースのアドインに関する同様の記事については、「[シングル サインオンを使用する ASP.NET Office アドインを作成する](create-sso-office-add-ins-aspnet.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="0cc95-106">For a similar article about an ASP.NET-based add-in, see [Create an ASP.NET Office Add-in that uses single sign-on](create-sso-office-add-ins-aspnet.md).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="0cc95-107">前提条件</span><span class="sxs-lookup"><span data-stu-id="0cc95-107">Prerequisites</span></span>

* <span data-ttu-id="0cc95-108">[Node および npm](https://nodejs.org/en/)、バージョン 6.9.4 以降</span><span class="sxs-lookup"><span data-stu-id="0cc95-108">[Node and npm](https://nodejs.org/en/), version 6.9.4 or later</span></span>

* <span data-ttu-id="0cc95-109">[Git バッシュ](https://git-scm.com/downloads) (またはその他の Git クライアント)</span><span class="sxs-lookup"><span data-stu-id="0cc95-109">[Git Bash](https://git-scm.com/downloads) (or another git client)</span></span>

* <span data-ttu-id="0cc95-110">TypeScript バージョン 2.2.2 以降</span><span class="sxs-lookup"><span data-stu-id="0cc95-110">TypeScript version 2.2.2 or later</span></span>

* <span data-ttu-id="0cc95-111">Office 365 (Office のサブスクリプション バージョン)。</span><span class="sxs-lookup"><span data-stu-id="0cc95-111">Office 365 (the subscription version of Office).</span></span> <span data-ttu-id="0cc95-112">Insider チャネルからの最新の月次バージョンとビルド。</span><span class="sxs-lookup"><span data-stu-id="0cc95-112">Latest monthly version and build from the Insiders channel.</span></span> <span data-ttu-id="0cc95-113">このバージョンを入手するには、Office Insider への参加が必要です。</span><span class="sxs-lookup"><span data-stu-id="0cc95-113">You need to be an Office Insider to get this version.</span></span> <span data-ttu-id="0cc95-114">詳細については、「[Office Insider になる](https://products.office.com/office-insider?tab=tab-1)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="0cc95-114">For more information, see [Be an Office Insider](https://products.office.com/office-insider?tab=tab-1).</span></span> <span data-ttu-id="0cc95-115">ビルドが半期チャネルの運用に移行すると、そのビルドで SSO を含むプレビュー機能のサポートはオフになりますので、ご注意ください。</span><span class="sxs-lookup"><span data-stu-id="0cc95-115">Please note that when a build graduates to the production semi-annual channel, support for preview features, including SSO, is turned off for that build.</span></span>

## <a name="set-up-the-starter-project"></a><span data-ttu-id="0cc95-116">スタート プロジェクトをセットアップする</span><span class="sxs-lookup"><span data-stu-id="0cc95-116">Set up the starter project</span></span>

1. <span data-ttu-id="0cc95-117">「[Office Add-in NodeJS SSO](https://github.com/officedev/office-add-in-nodejs-sso)」にあるリポジトリを複製するかダウンロードします。</span><span class="sxs-lookup"><span data-stu-id="0cc95-117">Clone or download the repo at [Office Add-in NodeJS SSO](https://github.com/officedev/office-add-in-nodejs-sso).</span></span>

    > [!NOTE]
    > <span data-ttu-id="0cc95-118">このサンプルには、次の 3 つのバージョンがあります。</span><span class="sxs-lookup"><span data-stu-id="0cc95-118">There are three versions of the sample:</span></span>  
    > * <span data-ttu-id="0cc95-p103">**[Before]** フォルダーはスタート プロジェクトです。SSO や承認に直接関連しない UI などの側面は、既に完了しています。この記事で後述する各セクションでは、これを完成させるための手順を順に説明します。</span><span class="sxs-lookup"><span data-stu-id="0cc95-p103">The **Before** folder is a starter project. The UI and other aspects of the add-in that are not directly connected to SSO or authorization are already done. Later sections of this article walk you through the process of completing it.</span></span>
    > * <span data-ttu-id="0cc95-p104">このサンプルの **[Completed]** バージョンは、この記事の手順を完了したときに得られるアドインと同様のものですが、完成済みのプロジェクトには、この記事のテキストと重複するコード コメントが含まれています。完成済みのバージョンを使用する場合は、この記事の手順をそのまま実行しますが、[Before] を [Completed] に置き換えて、「**クライアント側のコードを作成する**」と「**サーバー側のコードを作成する**」のセクションを省略してください。</span><span class="sxs-lookup"><span data-stu-id="0cc95-p104">The **Completed** version of the sample is just like the add-in that you would have if you completed the procedures of this article, except that the completed project has code comments that would be redundant with the text of this article. To use the completed version, just follow the instructions in this article, but replace "Before" with "Completed" and skip the sections **Code the client side** and **Code the server** side.</span></span>
    > * <span data-ttu-id="0cc95-124">**完成版のマルチテナント** バージョンは、マルチテナント機能をサポートする完成版のサンプルです。</span><span class="sxs-lookup"><span data-stu-id="0cc95-124">The **Completed Multitenant** version is a completed sample that supports multitenancy.</span></span> <span data-ttu-id="0cc95-125">SSO を使用する異なるドメインの Microsoft アカウントをサポートする場合は、このサンプルを確認してください。</span><span class="sxs-lookup"><span data-stu-id="0cc95-125">Explore this sample if you intend to support Microsoft accounts from different domains with SSO.</span></span>
    >
    > <span data-ttu-id="0cc95-126">_ローカル ホストの証明書は、使用するバージョンにかかわらず信頼する必要があります。リポジトリのリリース ノートの「IMPORTANT」 (重要) のメモを参照してください。_</span><span class="sxs-lookup"><span data-stu-id="0cc95-126">_Regardless of which version you use, you will need to trust a certificate for the localhost. See the "IMPORTANT" note in the Readme of the repo._</span></span>

1. <span data-ttu-id="0cc95-127">**[Before]** フォルダー内で Git bash コンソールを開きます。</span><span class="sxs-lookup"><span data-stu-id="0cc95-127">Open a Git bash console in the **Before** folder.</span></span>

1. <span data-ttu-id="0cc95-128">コンソールで `npm install` を入力して、package.json ファイル内のアイテム化されたすべての依存関係をインストールします。</span><span class="sxs-lookup"><span data-stu-id="0cc95-128">Enter `npm install` in the console to install all of the dependencies itemized in the package.json file.</span></span>

1. <span data-ttu-id="0cc95-129">コンソールで `npm run build` を入力して、プロジェクトをビルドします。</span><span class="sxs-lookup"><span data-stu-id="0cc95-129">Enter `npm run build` in the console to build the project.</span></span>

    > [!NOTE]
    > <span data-ttu-id="0cc95-p106">いくつかの使用されていない変数が宣言されているという、ビルド エラーが発生することがあります。これらのエラーは無視してください。これらは、後で追加する一部のコードが見つからないという「Before」バージョンのサンプルの副作用です。</span><span class="sxs-lookup"><span data-stu-id="0cc95-p106">You may see some build errors saying that some variables are declared but not used. Ignore these errors. They are a side effect of the fact that the "Before" version of the sample is missing some code that will be added later.</span></span>

## <a name="register-the-add-in-with-azure-ad-v20-endpoint"></a><span data-ttu-id="0cc95-133">Azure AD v2.0 エンドポイントにアドインを登録する</span><span class="sxs-lookup"><span data-stu-id="0cc95-133">Register the add-in with Azure AD v2.0 endpoint</span></span>

<span data-ttu-id="0cc95-134">次の手順は、複数の場所で使用できるように、一般的に記述されています。</span><span class="sxs-lookup"><span data-stu-id="0cc95-134">The following instruction are written generically so they can be used in multiple places.</span></span> <span data-ttu-id="0cc95-135">この記事では、次の手順を行います。</span><span class="sxs-lookup"><span data-stu-id="0cc95-135">For this article do the following:</span></span>

- <span data-ttu-id="0cc95-136">プレースホルダー **$ADD-IN-NAME$** を `Office-Add-in-NodeJS-SSO` に置き換えます。</span><span class="sxs-lookup"><span data-stu-id="0cc95-136">Replace the placeholder **$ADD-IN-NAME$** with `Office-Add-in-NodeJS-SSO`.</span></span>
- <span data-ttu-id="0cc95-137">プレースホルダー **$FQDN-WITHOUT-PROTOCOL$** を `localhost:3000` に置き換えます。</span><span class="sxs-lookup"><span data-stu-id="0cc95-137">Replace the placeholder **$FQDN-WITHOUT-PROTOCOL$** with `localhost:3000`.</span></span>
- <span data-ttu-id="0cc95-138">**[アクセス許可の選択]** ダイアログでアクセス許可を指定するときに、次のアクセス許可のチェック ボックスをオンにします。</span><span class="sxs-lookup"><span data-stu-id="0cc95-138">When you specify permissions in the **Select Permissions** dialog, check the boxes for the following permissions.</span></span> <span data-ttu-id="0cc95-139">アドイン自体に実際に必要なものは最初のもののみですが、Office ホストがアドインの Web アプリケーションへのトークンを取得するには、`profile` のアクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="0cc95-139">Only the first is really required by your add-in itself; but the `profile` permission is required for the Office host to get a token to your add-in web application.</span></span>
  * <span data-ttu-id="0cc95-140">Files.Read.All</span><span class="sxs-lookup"><span data-stu-id="0cc95-140">Files.Read.All</span></span>
  * <span data-ttu-id="0cc95-141">profile</span><span class="sxs-lookup"><span data-stu-id="0cc95-141">profile</span></span>

[!INCLUDE[](../includes/register-sso-add-in-aad-v2-include.md)]


## <a name="grant-administrator-consent-to-the-add-in"></a><span data-ttu-id="0cc95-142">アドインに管理者の同意を許可する</span><span class="sxs-lookup"><span data-stu-id="0cc95-142">Grant administrator consent to the add-in</span></span>

[!INCLUDE[](../includes/grant-admin-consent-to-an-add-in-include.md)]

## <a name="configure-the-add-in"></a><span data-ttu-id="0cc95-143">アドインを構成する</span><span class="sxs-lookup"><span data-stu-id="0cc95-143">Configure the add-in</span></span>

1. <span data-ttu-id="0cc95-p109">コード エディターで、src\server.ts ファイルを開きます。先頭近くに、`AuthModule` クラスのコンストラクターの呼び出しがあります。コンストラクターには、値を割り当てる必要がある、文字列のパラメーターがあります。</span><span class="sxs-lookup"><span data-stu-id="0cc95-p109">In your code editor, open the src\server.ts file. Near the top there is a call to a constructor of an `AuthModule` class. There are some string parameters in the constructor to which you need to assign values.</span></span>

1. <span data-ttu-id="0cc95-p110">`client_id` プロパティの場合は、アドインの登録時に保存したアプリケーション ID でプレースホルダーの `{client GUID}` を置き換えます。完了すると、単一引用符で囲まれた GUID のみになります。"{}" 文字は取り去る必要があります。</span><span class="sxs-lookup"><span data-stu-id="0cc95-p110">For the `client_id` property, replace the placeholder `{client GUID}` with the application ID that you saved when you registered the add-in. When you are done, there should just be a GUID in single quotation marks. There should not be any "{}" characters.</span></span>

1. <span data-ttu-id="0cc95-150">`client_secret` プロパティのプレースホルダー `{client secret}` は、アドインの登録時に保存したアプリケーション シークレットに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="0cc95-150">For the `client_secret` property, replace the placeholder `{client secret}` with the application secret that you saved when you registered the add-in.</span></span>

1. <span data-ttu-id="0cc95-p111">`audience` プロパティの場合は、アドインの登録時に保存したアプリケーション ID でプレースホルダーの `{audience GUID}` を置き換えます。(`client_id` プロパティに割り当てた値とまったく同じになります)。</span><span class="sxs-lookup"><span data-stu-id="0cc95-p111">For the `audience` property, replace the placeholder `{audience GUID}` with the application ID that you saved when you registered the add-in. (The very same value that you assigned to the `client_id` property.)</span></span>
  
1. <span data-ttu-id="0cc95-153">`issuer` プロパティに割り当てた文字列には、*{O365 tenant GUID}* のプレースホルダーがあります。</span><span class="sxs-lookup"><span data-stu-id="0cc95-153">In the string assigned to the `issuer` property, you will see the placeholder *{O365 tenant GUID}*.</span></span> <span data-ttu-id="0cc95-154">これを Office 365 のテナント ID に置き換えます。</span><span class="sxs-lookup"><span data-stu-id="0cc95-154">Replace this with the Office 365 tenancy ID.</span></span> <span data-ttu-id="0cc95-155">「[Office 365 のテナント ID を検索する](/onedrive/find-your-office-365-tenant-id)」に記載されているいずれかの方法で、テナント ID を取得します。</span><span class="sxs-lookup"><span data-stu-id="0cc95-155">Use one of the methods in [Find your Office 365 tenant ID](/onedrive/find-your-office-365-tenant-id) to obtain it.</span></span> <span data-ttu-id="0cc95-156">完了すると、`issuer` プロパティの値は、次のようになります。</span><span class="sxs-lookup"><span data-stu-id="0cc95-156">When you are done, the `issuer` property value should look something like this:</span></span>

    `https://login.microsoftonline.com/12345678-1234-1234-1234-123456789012/v2.0`

1. <span data-ttu-id="0cc95-p113">`AuthModule` コンストラクターのその他の値は未変更のままにしておきます。 ファイルを保存して閉じます。</span><span class="sxs-lookup"><span data-stu-id="0cc95-p113">Leave the other parameters in the `AuthModule` constructor unchanged. Save and close the file.</span></span>

1. <span data-ttu-id="0cc95-159">プロジェクトのルートにある、アドイン マニフェスト ファイル「Office-Add-in-NodeJS-SSO.xml」を開きます。</span><span class="sxs-lookup"><span data-stu-id="0cc95-159">In the root of the project, open the add-in manifest file “Office-Add-in-NodeJS-SSO.xml”.</span></span>

1. <span data-ttu-id="0cc95-160">ファイルの最後までスクロールします。</span><span class="sxs-lookup"><span data-stu-id="0cc95-160">Scroll to the bottom of the file.</span></span>

1. <span data-ttu-id="0cc95-161">最後の `</VersionOverrides>` タグの直前に、次に示すマークアップが見つかります。</span><span class="sxs-lookup"><span data-stu-id="0cc95-161">Just above the end `</VersionOverrides>` tag, you will find the following markup:</span></span>

    ```xml
    <WebApplicationInfo>
      <Id>{application_GUID here}</Id>
      <Resource>api://localhost:3000/{application_GUID here}</Resource>
      <Scopes>
          <Scope>Files.Read.All</Scope>
          <Scope>profile</Scope>
      </Scopes>
    </WebApplicationInfo>
    ```

1. <span data-ttu-id="0cc95-162">このマークアップ内のプレースホルダー “{application_GUID here}” の*両方の場所*を、アドインの登録時にコピーしたアプリケーション ID に置き換えます </span><span class="sxs-lookup"><span data-stu-id="0cc95-162">Replace the placeholder “{application_GUID here}” *in both places* in the markup with the Application ID that you copied when you registered your add-in.</span></span> <span data-ttu-id="0cc95-163">("{}" は ID の一部ではないので、これらを含めないでください。)。これは、web.config の ClientID と Audience に使用したものと同じ ID です。</span><span class="sxs-lookup"><span data-stu-id="0cc95-163">(The "{}" are not part of the ID, so don't include them.) This is the same ID you used in for the ClientID and Audience in the web.config.</span></span>

    > [!NOTE]
    > * <span data-ttu-id="0cc95-164">**リソース**の値は、アドインの登録に Web API プラットフォームを追加したときに設定した**アプリケーション ID URI** です。</span><span class="sxs-lookup"><span data-stu-id="0cc95-164">The **Resource** value is the **Application ID URI** you set when you added the Web API platform to the registration of the add-in.</span></span>
    > * <span data-ttu-id="0cc95-165">**[範囲]** セクションは、アドインが AppSource から販売された場合に、同意ダイアログ ボックスを生成するためにのみ使用します。</span><span class="sxs-lookup"><span data-stu-id="0cc95-165">The **Scopes** section is used only to generate a consent dialog box if the add-in is sold through AppSource.</span></span>

1. <span data-ttu-id="0cc95-166">ファイルを保存して閉じます。</span><span class="sxs-lookup"><span data-stu-id="0cc95-166">Save and close the file.</span></span>

## <a name="code-the-client-side"></a><span data-ttu-id="0cc95-167">クライアント側のコードを作成する</span><span class="sxs-lookup"><span data-stu-id="0cc95-167">Code the client side</span></span>

1. <span data-ttu-id="0cc95-p115">**[public]** フォルダー内の program.js ファイルを開きます。これには、一部のコードが既に含まれています。</span><span class="sxs-lookup"><span data-stu-id="0cc95-p115">Open the program.js file in the **public** folder. It already has some code in it:</span></span>

    * <span data-ttu-id="0cc95-170">`Office.initialize` メソッドへの割り当てが、`getGraphAccessTokenButton` ボタン クリック イベントへのハンドラーの割り当てになります。</span><span class="sxs-lookup"><span data-stu-id="0cc95-170">An assignment to the `Office.initialize` method that, in turn, assigns a handler to the `getGraphAccessTokenButton` button click event.</span></span>
    * <span data-ttu-id="0cc95-171">`showResult` メソッドは、作業ウィンドウの下側に Microsoft Graph から返されたデータ (またはエラー メッセージ) を表示するものです。</span><span class="sxs-lookup"><span data-stu-id="0cc95-171">A `showResult` method that will display data returned from Microsoft Graph (or an error message) at the bottom of the task pane.</span></span>
    * <span data-ttu-id="0cc95-172">`logErrors` メソッドは、エンド ユーザーを対象としていないエラーをコンソールにログ出力するものです。</span><span class="sxs-lookup"><span data-stu-id="0cc95-172">A `logErrors` method that will log to console errors that are not intended for the end user.</span></span>

1. <span data-ttu-id="0cc95-p116">`Office.initialize` への割り当ての下に、次に示すコードを追加します。このコードについては、次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="0cc95-p116">Below the assignment to `Office.initialize`, add the code below. Note the following about this code:</span></span>

    * <span data-ttu-id="0cc95-p117">アドインのエラー処理により、アクセス トークンの取得が別のオプションのセットを使用して自動的に再試行されることがあります。 カウンター変数 `timesGetOneDriveFilesHasRun` とフラグ変数 `triedWithoutForceConsent` および `timesMSGraphErrorReceived` を使用して、失敗するトークン取得の繰り返しからユーザーが抜け出せるようにします。</span><span class="sxs-lookup"><span data-stu-id="0cc95-p117">The error-handling in the add-in will sometimes automatically attempt a second time to get an access token, using a different set of options. The counter variable `timesGetOneDriveFilesHasRun`, and the flag variables `triedWithoutForceConsent` and `timesMSGraphErrorReceived` are used to ensure that the user isn't cycled repeatedly through failed attempts to get a token.</span></span>
    * <span data-ttu-id="0cc95-p118">この後の手順では `getDataWithToken` メソッドを作成しますが、そのメソッドで `forceConsent` というオプションが `false` に設定される点に注意してください。詳細については、次の手順で説明します。</span><span class="sxs-lookup"><span data-stu-id="0cc95-p118">You create the `getDataWithToken` method in the next step, but note that it sets an option called `forceConsent` to `false`. More about that in the next step.</span></span>

    ```javascript
    var timesGetOneDriveFilesHasRun = 0;
    var triedWithoutForceConsent = false;
    var timesMSGraphErrorReceived = false;

    function getOneDriveFiles() {
        timesGetOneDriveFilesHasRun++;
        triedWithoutForceConsent = true;
        getDataWithToken({ forceConsent: false });
    }
    ```

1. <span data-ttu-id="0cc95-p119">`getOneDriveFiles` メソッドの下に、次のコードを追加します。このコードについては、次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="0cc95-p119">Below the `getOneDriveFiles` method, add the code below. Note the following about this code:</span></span>

    * <span data-ttu-id="0cc95-181">[getAccessTokenAsync](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) は Office.js の新しい API です。これにより、アドインは Office ホスト アプリケーション (Excel、PowerPoint、Word など) に、アドインへの (Office にサインインしているユーザーの) アクセス トークンを要求できるようになります。</span><span class="sxs-lookup"><span data-stu-id="0cc95-181">The [getAccessTokenAsync](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) is the new API in Office.js that enables an add-in to ask the Office host application (Excel, PowerPoint, Word, etc.) for an access token to the add-in (for the user signed into Office).</span></span> <span data-ttu-id="0cc95-182">その結果、この Office ホスト アプリケーションによって、Azure AD 2.0 エンドポイントにこのトークンが要求されます。</span><span class="sxs-lookup"><span data-stu-id="0cc95-182">The Office host application, in turn, asks the Azure AD 2.0 endpoint for the token.</span></span> <span data-ttu-id="0cc95-183">アドインの登録時に、アドインに対する Office ホストを事前認証しているため、Azure AD はそのトークンを送信します。</span><span class="sxs-lookup"><span data-stu-id="0cc95-183">Since you preauthorized the Office host to your add-in when you registered it, Azure AD will send the token.</span></span>
    * <span data-ttu-id="0cc95-184">Office にサインインしているユーザーがいない場合、Office ホストはユーザーにサインインを求めるダイアログを表示します。</span><span class="sxs-lookup"><span data-stu-id="0cc95-184">If no user is signed into Office, the Office host will prompt the user to sign in.</span></span>
    * <span data-ttu-id="0cc95-p121">オプションのパラメーター `forceConsent` を `false` に設定すると、ユーザーがアドインを使用するたびに、Office ホストにアドインへのアクセス権を付与するための同意を求めるダイアログが表示されなくなります。 ユーザーが初めてアドインを実行すると、`getAccessTokenAsync` の呼び出しは失敗しますが、この後の手順で追加するエラー処理ロジックにより、`forceConsent` オプションを `true` に設定した再呼び出しが自動的に実行され、ユーザーに同意を求めるダイアログが表示されます。ただし、これは初回時のみ実行されます。</span><span class="sxs-lookup"><span data-stu-id="0cc95-p121">The options parameter sets `forceConsent` to `false`, so the user will not be prompted to consent to giving the Office host access to your add-in every time she or he uses the add-in. The first time the user runs the add-in, the call of `getAccessTokenAsync` will fail, but error-handling logic that you add in a later step will automatically re-call with the `forceConsent` option set to `true` and the user will be prompted to consent, but only that first time.</span></span>
    * <span data-ttu-id="0cc95-187">`handleClientSideErrors` メソッドは、この後の手順で作成します。</span><span class="sxs-lookup"><span data-stu-id="0cc95-187">You will create the `handleClientSideErrors` method in a later step.</span></span>

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

1. <span data-ttu-id="0cc95-p122">TODO1 を次に示す行に置き換えます。`getData` メソッドとサーバー側の "/api/values" ルートは、この後の手順で作成します。エンドポイントには、相対 URL を使用します。これは、その URL がアドインと同じドメインでホストされている必要があるためです。</span><span class="sxs-lookup"><span data-stu-id="0cc95-p122">Replace the TODO1 with the following lines. You create the `getData` method and the server-side “/api/values” route in later steps. A relative URL is used for the endpoint because it must be hosted on the same domain as your add-in.</span></span>

    ```javascript
    accessToken = result.value;
    getData("/api/values", accessToken);
    ```

1. <span data-ttu-id="0cc95-p123">`getOneDriveFiles` メソッドの下に、以下を追加します。このコードについては、次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="0cc95-p123">Below the `getOneDriveFiles` method, add the following. About this code, note:</span></span>

    * <span data-ttu-id="0cc95-p124">このメソッドは、特定の Web API エンドポイントを呼び出して、Office ホスト アプリケーションがアドインへのアクセスに使用したものと同じアクセス トークンを渡します。サーバー側では、このアクセス トークンが Microsoft Graph へのアクセス トークンを取得するための「代理 (on-behalf-of)」フローで使用されます。</span><span class="sxs-lookup"><span data-stu-id="0cc95-p124">This method calls a specified Web API endpoint and passes it the same access token that the Office host application used to get access to your add-in. On the server-side, this access token will be used in the “on behalf of” flow to obtain an access token to Microsoft Graph.</span></span>
    * <span data-ttu-id="0cc95-195">`handleServerSideErrors` メソッドは、この後の手順で作成します。</span><span class="sxs-lookup"><span data-stu-id="0cc95-195">You will create the `handleServerSideErrors` method in a later step.</span></span>

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

### <a name="create-the-error-handling-methods"></a><span data-ttu-id="0cc95-196">エラー処理のメソッドを作成する</span><span class="sxs-lookup"><span data-stu-id="0cc95-196">Create the error-handling methods</span></span>

1. <span data-ttu-id="0cc95-p125">`getData` メソッドの下に、次のメソッドを追加します。 このメソッドは、Office ホストがアドインの Web サービスへのアクセス トークンを取得できないときに、アドインのクライアントでエラーを処理します。 こうしたエラーはエラー コードで報告されるため、このメソッドでは `switch` ステートメントを使用してエラーを識別します。</span><span class="sxs-lookup"><span data-stu-id="0cc95-p125">Below the `getData` method, add the following method. This method will handle errors in the add-in's client when the Office host is unable to obtain an access token to the add-in's web service. These errors are reported with an error code, so the method uses a `switch` statement to distinguish them.</span></span>

    ```javascript
    function handleClientSideErrors(result) {

        switch (result.error.code) {

            // TODO2: Handle the case where user is not logged in, or the user cancelled, without responding, a
            //        prompt to provide a 2nd authentication factor.

            // TODO3: Handle the case where the user's sign-in or consent was aborted.

            // TODO4: Handle the case where the user is logged in with an account that is neither work or school,
            //        nor Microsoft Account.

            // TODO5: Handle the case where the Office host has not been authorized to the add-in's web service or
            //        the user has not granted the service permission to their `profile`.

            // TODO6: Handle an unspecified error from the Office host.

            // TODO7: Handle the case where the Office host cannot get an access token to the add-ins
            //        web service/application.

            // TODO8: Handle the case where the user triggered an operation that calls `getAccessTokenAsync`
            //        before a previous call of it completed.

            // TODO9: Handle the case where the add-in does not support forcing consent.

            // TODO10: Log all other client errors.
        }
    }
    ```

1. <span data-ttu-id="0cc95-p126">`TODO2` を次のコードに置き換えます。 エラー 13001 は、ユーザーがログインしていない場合、または 2 番目の認証要素の指定を求めるダイアログに応答しないでキャンセルした場合に発生します。 どちらの場合も、このコードでは `getDataWithToken` メソッドを再実行して、サインインを求めるダイアログの表示を強制するようにオプションを設定します。</span><span class="sxs-lookup"><span data-stu-id="0cc95-p126">Replace `TODO2` with the following code. Error 13001 occurs when the user is not logged in, or the user cancelled, without responding, a prompt to provide a 2nd authentication factor. In either case, the code re-runs the `getDataWithToken` method and sets an option to force a sign-in prompt.</span></span>

    ```javascript
    case 13001:
        getDataWithToken({ forceAddAccount: true });
        break;
    ```

1. <span data-ttu-id="0cc95-p127">`TODO3` を次のコードに置き換えます。 エラー 13002 は、ユーザーのサインインまたは同意が中断された場合に発生します。 ユーザーに対して 1 回だけ再試行を求めます。</span><span class="sxs-lookup"><span data-stu-id="0cc95-p127">Replace `TODO3` with the following code. Error 13002 occurs when user's sign-in or consent was aborted. Ask the user to try again but no more than once again.</span></span>

    ```javascript
    case 13002:
        if (timesGetOneDriveFilesHasRun < 2) {
            showResult(['Your sign-in or consent was aborted before completion. Please try that operation again.']);
        } else {
            logError(result);
        }
        break;
    ```

1. <span data-ttu-id="0cc95-206">`TODO4` を次のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="0cc95-206">Replace `TODO4` with the following code.</span></span> <span data-ttu-id="0cc95-207">エラー 13003 は、ユーザーが職場または学校アカウント、または Microsoft アカウントのいずれでもないアカウントでログインしている場合に発生します。</span><span class="sxs-lookup"><span data-stu-id="0cc95-207">Error 13003 occurs when user is logged in with an account that is neither work or school, nor Microsoft Account.</span></span> <span data-ttu-id="0cc95-208">サインアウト後にサポートされているアカウントの種類でもう一度サインインするよう、ユーザーに求めます。</span><span class="sxs-lookup"><span data-stu-id="0cc95-208">Ask the user to sign-out and then in again with a supported account type.</span></span>

    ```javascript
    case 13003:
        showResult(['Please sign out of Office and sign in again with a work or school account, or Microsoft Account. Other kinds of accounts, like corporate domain accounts do not work.']);
        break;
    ```

    > [!NOTE]
    > <span data-ttu-id="0cc95-209">エラー 13004 は開発時にのみ発生するため、このメソッドでは処理しません。</span><span class="sxs-lookup"><span data-stu-id="0cc95-209">Errors 13004 and 13005 are not handled in this method because they should only occur in development.</span></span> <span data-ttu-id="0cc95-210">これはランタイム コードで修正できるものではなく、エンド ユーザーに報告しても意味がありません。</span><span class="sxs-lookup"><span data-stu-id="0cc95-210">They cannot be fixed by runtime code and there would be no point in reporting them to an end user.</span></span>

1. <span data-ttu-id="0cc95-211">`TODO5` を次のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="0cc95-211">Replace `TODO5` with the following code.</span></span> <span data-ttu-id="0cc95-212">Office がアドインの Web サービスに承認されていない場合や、ユーザーの `profile` に対してサービスのアクセス許可が付与されていない場合、エラー 13005 が発生します。</span><span class="sxs-lookup"><span data-stu-id="0cc95-212">Error 13005 occurs when Office has not been authorized to the add-in's web service or the user has not granted the service permission to their `profile`.</span></span>

    ```javascript
    case 13005:
        getDataWithToken({ forceConsent: true });
        break;
    ```

1. <span data-ttu-id="0cc95-p131">`TODO6` を次のコードと置き換えます。エラー 13006 は、Office ホストで未指定のエラーがある場合に発生します。ホストが不安定な状態にあることを示している可能性があります。ユーザーに Office の再起動を求めます。</span><span class="sxs-lookup"><span data-stu-id="0cc95-p131">Replace `TODO6` with the following code. Error 13006 occurs when there has been an unspecified error in the Office host that may indicate that the host is in an unstable state. Ask the user to restart Office.</span></span>

    ```javascript
    case 13006:
        showResult(['Please save your work, sign out of Office, close all Office applications, and restart this Office application.']);
        break;
    ```

1. <span data-ttu-id="0cc95-p132">`TODO7` を次のコードに置き換えます。 エラー 13007 は、Office ホストの AAD との相互作用に問題があり、ホストがアドイン Web サービス/アプリケーションへのアクセス トークンを取得できない場合に発生します。 ネットワークに一時的な問題が発生している可能性があります。 しばらく待ってから再試行するようにユーザーに求めます。</span><span class="sxs-lookup"><span data-stu-id="0cc95-p132">Replace `TODO7` with the following code. Error 13007 occurs when something has gone wrong with the Office host's interaction with AAD so the host cannot get an access token to the add-ins web service/application. This may be a temporary network issue. Ask the user to try again later.</span></span>

    ```javascript
    case 13007:
        showResult(['That operation cannot be done at this time. Please try again later.']);
        break;
    ```

1. <span data-ttu-id="0cc95-220">`TODO8` を次のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="0cc95-220">Replace `TODO8` with the following code.</span></span> <span data-ttu-id="0cc95-221">エラー 13008 は、前回の `getAccessTokenAsync` の呼び出しが完了する前に、それを呼び出す操作をユーザーがトリガーしたときに発生します。</span><span class="sxs-lookup"><span data-stu-id="0cc95-221">Error 13008 occurs when the user triggered an operation that calls `getAccessTokenAsync` before a previous call of it completed.</span></span>

    ```javascript
    case 13008:
        showResult(['Please try that operation again after the current operation has finished.']);
        break;
    ```

1. <span data-ttu-id="0cc95-p134">`TODO9` を次のコードに置き換えます。 エラー 13009 は、アドインが強制的な同意をサポートしていないときに、`getAccessTokenAsync` オプションを `forceConsent` に設定して `true` を呼び出した場合に発生します。 通常、この場合は、コードによって同意オプションを `getAccessTokenAsync` に設定して自動的に `false` を再実行する必要があります。 ただし、`forceConsent` を `true` に設定してメソッドを呼び出すこと自体が、そのオプションを `false` に設定したメソッドの呼び出しで発生したエラーに対する自動的な応答の場合もあります。 その場合は、コードで再試行するのではなく、ユーザーにサインアウトしてから再度サインインするように通知する必要があります。</span><span class="sxs-lookup"><span data-stu-id="0cc95-p134">Replace `TODO9` with the following code. Error 13009 occurs when the add-in does not support forcing consent, but `getAccessTokenAsync` was called with the `forceConsent` option set to `true`. In the usual case when this happens the code should automatically re-run `getAccessTokenAsync` with the consent option set to `false`. However, in some cases, calling the method with `forceConsent` set to `true` was itself an automatic response to an error in a call to the method with the option set to `false`. In that case, the code should not try again, but instead it should advise the user to sign out and sign in again.</span></span>

    ```javascript
    case 13009:
        if (triedWithoutForceConsent) {
            showResult(['Please sign out of Office and sign in again with a work or school account, or Microsoft Account.']);
        } else {
            getDataWithToken({ forceConsent: false });
        }
        break;


1. Replace `TODO10` with the following code.

    ```javascript
    default:
        logError(result);
        break;
    ```  

1. <span data-ttu-id="0cc95-p135">`handleClientSideErrors` メソッドの下に、次のメソッドを追加します。このメソッドは、代理 (on-behalf-of) フローの実行時または Microsoft Graph からのデータの取得時の問題により、アドインの Web サービスで発生したエラーを処理します。</span><span class="sxs-lookup"><span data-stu-id="0cc95-p135">Below the `handleClientSideErrors` method, add the following method. This method will handle errors in the add-in's web service when something goes wrong in executing the on-behalf-of flow or in getting data from Microsoft Graph.</span></span>

    ```javascript
    function handleServerSideErrors(result) {

        // TODO11: Handle the case where AAD asks for an additional form of authentication.

        // TODO12: Handle the case where consent has not been granted, or has been revoked.

        // TODO13: Handle the case where an invalid scope (permission) was used in the on-behalf-of flow

        // TODO14: Handle the case where the token that the add-in's client-side sends to its
        //         server-side is not valid because it is missing `access_as_user` scope (permission).

        // TODO15: Handle the case where the token sent to Microsoft Graph in the request for
        //         data is expired or invalid.

        // TODO16: Log all other server errors.
    }
    ```

1. <span data-ttu-id="0cc95-229">`TODO11` を次のコードに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="0cc95-229">Replace `TODO11` with the following code.</span></span> <span data-ttu-id="0cc95-230">このコードの注意点は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="0cc95-230">Note about this code:</span></span>

    * <span data-ttu-id="0cc95-p137">ユーザーがパスワードだけで Office にサインオンできる場合でも、Microsoft Graph のいくつかのターゲット (たとえば、OneDrive) にアクセスするために、追加の認証要素を提供するようにユーザーに要求する、Azure Active Directory の構成があります。その場合、AAD は `Claims` プロパティを含むエラー 50076 で応答を送信します。</span><span class="sxs-lookup"><span data-stu-id="0cc95-p137">There are configurations of Azure Active Directory in which the user is required to provide additional authentication factor(s) to access some Microsoft Graph targets (e.g., OneDrive), even if the user can sign on to Office with just a password. In that case, AAD will send a response, with error 50076, that has a `Claims` property.</span></span>
    * <span data-ttu-id="0cc95-p138">Office ホストは、\*\*\*\* オプションとして `authChallenge` 値を使用して新しいトークンを取得します。 これにより、認証のすべての必要なフォームをユーザーに表示するように AAD に指示します。</span><span class="sxs-lookup"><span data-stu-id="0cc95-p138">The Office host should get a new token with the **Claims** value as the `authChallenge` option. This tells AAD to prompt the user for all required forms of authentication.</span></span>

    ```javascript
    if (result.responseJSON.error.innerError
            && result.responseJSON.error.innerError.error_codes
            && result.responseJSON.error.innerError.error_codes[0] === 50076){
        getDataWithToken({ authChallenge: result.responseJSON.error.innerError.claims });
    }
    ```

1. <span data-ttu-id="0cc95-p139">`TODO12` を次のコードに置き換えます (*前の手順で追加したコードの最後にある右波かっこのすぐ下*)。このコードの注意点は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="0cc95-p139">Replace `TODO12` with the following code *just below the last closing brace of the code you added in the previous step*. Note about this code:</span></span>

    * <span data-ttu-id="0cc95-237">エラー 65001 は、1 つ以上のアクセス許可について Microsoft Graph にアクセスするための同意が与えられていない (または取り消されている) ことを意味します。</span><span class="sxs-lookup"><span data-stu-id="0cc95-237">Error 65001 means that consent to access Microsoft Graph was not granted (or was revoked) for one or more permissions.</span></span>
    * <span data-ttu-id="0cc95-238">アドインでは、`forceConsent` オプションを `true` に設定して新しいトークンを取得する必要があります。</span><span class="sxs-lookup"><span data-stu-id="0cc95-238">The add-in should get a new token with the `forceConsent` option set to `true`.</span></span>

    ```javascript
    else if (result.responseJSON.error.innerError
            && result.responseJSON.error.innerError.error_codes
            && result.responseJSON.error.innerError.error_codes[0] === 65001){
        getDataWithToken({ forceConsent: true });
    }
    ```

1. <span data-ttu-id="0cc95-p140">`TODO13` を次のコードに置き換えます (*前の手順で追加したコードの最後にある右波かっこのすぐ下*)。このコードの注意点は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="0cc95-p140">Replace `TODO13` with the following code *just below the last closing brace of the code you added in the previous step*. Note about this code:</span></span>

    * <span data-ttu-id="0cc95-p141">エラー 70011 は、無効なスコープ (アクセス許可) が要求されたことを示します。 アドインでは、エラーを報告する必要があります。</span><span class="sxs-lookup"><span data-stu-id="0cc95-p141">Error 70011 means that an invalid scope (permission) has been requested. The add-in should report the error.</span></span>
    * <span data-ttu-id="0cc95-243">コードでは、その他のエラーを AAD エラー番号と共に記録します。</span><span class="sxs-lookup"><span data-stu-id="0cc95-243">The code logs any other error with an AAD error number.</span></span>

    ```javascript
    else if (result.responseJSON.error.innerError
            && result.responseJSON.error.innerError.error_codes
            && result.responseJSON.error.innerError.error_codes[0] === 70011){
        showResult(['The add-in is asking for a type of permission that is not recognized.']);
    }
    ```

1. <span data-ttu-id="0cc95-p142">`TODO14` を次のコードに置き換えます (*前の手順で追加したコードの最後にある右波かっこのすぐ下*)。このコードの注意点は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="0cc95-p142">Replace `TODO14` with the following code *just below the last closing brace of the code you added in the previous step*. Note about this code:</span></span>

    * <span data-ttu-id="0cc95-246">この後の手順で作成するサーバー側のコードでは、アドインのクライアントが AAD に送信して代理 (on-behalf-of) フローで使用されるアクセス トークンに `access_as_user` スコープ (アクセス許可) が含まれていない場合に、末尾が `... expected access_as_user` のメッセージを送信します。</span><span class="sxs-lookup"><span data-stu-id="0cc95-246">Server-side code that you create in a later step will send the message that ends with `... expected access_as_user` if the `access_as_user` scope (permission) is not in the access token that the add-in's client sends to AAD to be used in the on-behalf-of flow.</span></span>
    * <span data-ttu-id="0cc95-247">アドインでは、エラーを報告する必要があります。</span><span class="sxs-lookup"><span data-stu-id="0cc95-247">The add-in should report the error.</span></span>

    ```javascript
    else if (result.responseJSON.error.name
            && result.responseJSON.error.name.indexOf('expected access_as_user') !== -1){
        showResult(['Microsoft Office does not have permission to get Microsoft Graph data on behalf of the current user.']);
    }
    ```

1. <span data-ttu-id="0cc95-p143">`TODO15` を次のコードに置き換えます (*前の手順で追加したコードの最後にある右波かっこのすぐ下*)。このコードの注意点は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="0cc95-p143">Replace `TODO15` with the following code *just below the last closing brace of the code you added in the previous step*. Note about this code:</span></span>

    * <span data-ttu-id="0cc95-250">有効期限切れのトークンや無効なトークンが Microsoft Graph に送信される可能性はほとんどありませんが、そのような事態が発生した場合は、この後の手順で作成するサーバー側のコードは、文字列 `Microsoft Graph error` で終了します。</span><span class="sxs-lookup"><span data-stu-id="0cc95-250">It is unlikely that an expired or invalid token will be sent to Microsoft Graph; but if it does happen, the server-side code that you will create in a later step will end with the string `Microsoft Graph error`.</span></span>
    * <span data-ttu-id="0cc95-p144">この場合、アドインは `timesGetOneDriveFilesHasRun` カウンター変数と `timesGetOneDriveFilesHasRun` フラグ変数をリセットしてから、ボタン ハンドラー メソッドを再呼び出しすることで、認証プロセス全体を最初から開始する必要があります。 ただし、これは 1 回のみ実行する必要があります。 この事態が再度発生した場合は、単にエラーを記録するようにします。</span><span class="sxs-lookup"><span data-stu-id="0cc95-p144">In this case, the add-in should start the entire authentication process over by resetting the `timesGetOneDriveFilesHasRun` counter and `timesGetOneDriveFilesHasRun` flag variables, and then re-calling the button handler method. But it should do this only once. If it happens again, it should just log the error.</span></span>
    * <span data-ttu-id="0cc95-254">コードでは、この事態が連続して 2 回発生した場合にエラーを記録します。</span><span class="sxs-lookup"><span data-stu-id="0cc95-254">The code logs the error if it happens twice in succession.</span></span>

    ```javascript
    else if (result.responseJSON.error.name
            && result.responseJSON.error.name.indexOf('Microsoft Graph error') !== -1) {
        if (!timesMSGraphErrorReceived) {
            timesMSGraphErrorReceived = true;
            timesGetOneDriveFilesHasRun = 0;
            triedWithoutForceConsent = false;
            getOneDriveFiles();
        } else {
            logError(result);
        }
    }
    ```

1. <span data-ttu-id="0cc95-255">`TODO16` を次のコードに置き換えます (*前の手順で追加したコードの最後にある右波かっこのすぐ下*)。</span><span class="sxs-lookup"><span data-stu-id="0cc95-255">Replace `TODO16` with the following code *just below the last closing brace of the code you added in the previous step*.</span></span>

    ```javascript
    else {
        logError(result);
    }
    ```

## <a name="code-the-server-side"></a><span data-ttu-id="0cc95-256">サーバー側のコードを作成する</span><span class="sxs-lookup"><span data-stu-id="0cc95-256">Code the server side</span></span>

<span data-ttu-id="0cc95-257">変更の必要があるサーバー側のファイルは 2 つあります。</span><span class="sxs-lookup"><span data-stu-id="0cc95-257">There are two server-side files that need to be modified.</span></span>

- <span data-ttu-id="0cc95-p145">src\auth.js では、承認のヘルパー関数を提供します。これには、各種の承認フローで使用される汎用のメンバーが既に含まれています。これには、「代理」フローを実装するための関数を追加する必要があります。</span><span class="sxs-lookup"><span data-stu-id="0cc95-p145">The src\auth.js provides authorization helper functions. It already has generic members that are used in a variety of authorization flows. We need to add functions to it that implement the "on behalf of" flow.</span></span>
- <span data-ttu-id="0cc95-p146">src\server.js ファイルには、サーバーと express ミドルウェアを実行するために必要な基本的なメンバーが含まれています。これには、ホーム ページと Microsoft Graph データを取得するための Web API を提供する関数を追加する必要があります。</span><span class="sxs-lookup"><span data-stu-id="0cc95-p146">The src\server.js file has the basic members need to run a server and express middleware. We need to add functions to it that serve the home page and a Web API for obtaining Microsoft Graph data.</span></span>

### <a name="create-a-method-to-exchange-tokens"></a><span data-ttu-id="0cc95-263">トークンを交換するためのメソッドを作成する</span><span class="sxs-lookup"><span data-stu-id="0cc95-263">Create a method to exchange tokens</span></span>

1. <span data-ttu-id="0cc95-p147">\src\auth.ts ファイルを開きます。`AuthModule` クラスに、次に示すメソッドを追加します。このコードについては、次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="0cc95-p147">Open the \src\auth.ts file. Add the method below to the `AuthModule` class. Note the following about this code:</span></span>

    * <span data-ttu-id="0cc95-p148">`jwt` パラメーターは、アプリケーションへのアクセス トークンです。「代理 (on-behalf-of)」フローでは、これはリソースへのアクセス トークンの AAD と交換されます。</span><span class="sxs-lookup"><span data-stu-id="0cc95-p148">The `jwt` parameter is the access token to the application. In the "on behalf of" flow, it is exchanged with AAD for an access token to the resource.</span></span>
    * <span data-ttu-id="0cc95-269">scopes パラメーターには既定の値がありますが、このサンプルではコード呼び出しによってオーバーライドされています。</span><span class="sxs-lookup"><span data-stu-id="0cc95-269">The scopes parameter has a default value, but in this sample it will be overridden by the calling code.</span></span>
    * <span data-ttu-id="0cc95-270">resource パラメーターは省略可能です。</span><span class="sxs-lookup"><span data-stu-id="0cc95-270">The resource parameter is optional.</span></span> <span data-ttu-id="0cc95-271">[Secure Token Service (STS)](/previous-versions/windows-identity-foundation/ee748490(v=msdn.10)) が AAD V 2.0 エンドポイントの場合は使用しないでください。</span><span class="sxs-lookup"><span data-stu-id="0cc95-271">It should not be used when the [Secure Token Service (STS)](/previous-versions/windows-identity-foundation/ee748490(v=msdn.10)) is the AAD V 2.0 endpoint.</span></span> <span data-ttu-id="0cc95-272">V 2.0 エンドポイントでは scopes から resource を推測し、resource が HTTP 要求で送信される場合に、エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="0cc95-272">The V 2.0 endpoint infers the resource from the scopes and it returns an error if a resource is sent in the HTTP Request.</span></span>
    * <span data-ttu-id="0cc95-p150">`catch` ブロック内で例外をスローしても、即時の "500 Internal Server Error" がクライアントに送信されることは*ありません*。 server.js ファイルでコードを呼び出すことで、この例外をキャッチしてから、その例外をクライアントに送信するエラー メッセージに変換します。</span><span class="sxs-lookup"><span data-stu-id="0cc95-p150">Throwing an exception in the `catch` block will *not* cause an immediate "500 Internal Server Error" to be sent to the client. Calling code in the server.js file will catch this exception and turn it into an error message that is sent to the client.</span></span>

        ```typescript
        private async exchangeForToken(jwt: string, scopes: string[] = ['openid'], resource?: string) {
            try {
                // TODO3: Construct the parameters that will be sent in the body of the
                //        HTTP Request to the STS that starts the "on behalf of" flow.
                // TODO4: Send the request to the STS.
                // TODO5: Catch errors from the STS and relay them to the client.
                // TODO6: Process the response and persist the access token to resource.
            }
            catch (exception) {
                throw new UnauthorizedError('Unable to obtain an access token to the resource'
                                            + JSON.stringify(exception),
                                            exception);
            }
        }
        ```

1. <span data-ttu-id="0cc95-p151">`TODO3` を次のコードに置き換えます。このコードの注意点は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="0cc95-p151">Replace `TODO3` with the following code. About this code, note:</span></span>
    * <span data-ttu-id="0cc95-p152">「代理」ワークフローをサポートする STS は、HTTP 要求の本文に特定のプロパティ/値ペアが含まれていることを期待します。このコードは、要求の本文になるオブジェクトを構築します。</span><span class="sxs-lookup"><span data-stu-id="0cc95-p152">An STS that supports the "on behalf of" flow expects certain property/value pairs in the body of the HTTP request. This code constructs an object that will become the body of the request.</span></span>
    * <span data-ttu-id="0cc95-279">resource プロパティは、リソースがメソッドに渡された場合にのみ本文に追加されます。</span><span class="sxs-lookup"><span data-stu-id="0cc95-279">A resource property is added to the body if, and only if, a resource was passed to the method.</span></span>

        ```typescript
        const v2Params = {
                client_id: this.clientId,
                client_secret: this.clientSecret,
                grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
                assertion: jwt,
                requested_token_use: 'on_behalf_of',
                scope: scopes.join(' ')
            };
            let finalParams = {};
            if (resource) {
                // In JavaScript we could just add the resource property to the v2Params
                // object, but that won't compile in TypeScript.
                let v1Params  = { resource: resource };  
                for(var key in v2Params) { v1Params[key] = v2Params[key]; }
                finalParams = v1Params;
            } else {
                finalParams = v2Params;
            }
        ```

1. <span data-ttu-id="0cc95-280">`TODO4` を次に示すコードに置き換えます。このコードでは、HTTP 要求を STS のトークン エンドポイントに送信します。</span><span class="sxs-lookup"><span data-stu-id="0cc95-280">Replace `TODO4` with the following code which sends the HTTP request to the token endpoint of the STS.</span></span>

    ```typescript
    const res = await fetch(`${this.stsDomain}/${this.tenant}/${this.tokenURLsegment}`, {
        method: 'POST',
        body: form(finalParams),
        headers: {
            'Accept': 'application/json',
            'Content-Type': 'application/x-www-form-urlencoded'
        }
    });
    ```

1. <span data-ttu-id="0cc95-p153">`TODO5` を次のコードに置き換えます。 例外をスローしても、即時の "500 Internal Server Error" がクライアントに送信*されない*点に注意してください。 server.js ファイルでコードを呼び出すことで、この例外をキャッチしてから、その例外をクライアントに送信するエラー メッセージに変換します。</span><span class="sxs-lookup"><span data-stu-id="0cc95-p153">Replace `TODO5` with the following code. Note that throwing an exception will *not* cause an immediate "500 Internal Server Error" to be sent to the client. Calling code in the server.js file will catch this exception and turn it into an error message that is sent to the client.</span></span>

    ```typescript
     if (res.status !== 200) {
        const exception = await res.json();
        throw exception;
    }
    ```

1. <span data-ttu-id="0cc95-p154">`TODO6` を次に示すコードに置き換えます。このコードはリソースへのアクセス トークンを永続化して、有効期限になると、そのアクセス トークンを返します。コードを呼び出すことで、期限切れになっていないリソースへのアクセス トークンが再使用されるため、STS への不要な呼び出しを回避できます。この動作のしくみは、次のセクションで説明します。</span><span class="sxs-lookup"><span data-stu-id="0cc95-p154">Replace `TODO6` with the following code. Note that the code persists the access token to the resource, and it's expiration time, in addition to returning it. Calling code can avoid unnecessary calls to the STS by reusing an unexpired access token to the resource. You'll see how to do that in the next section.</span></span>

    ```typescript  
    const json = await res.json();
    const resourceToken = json['access_token'];
    ServerStorage.persist('ResourceToken', resourceToken);
    const expiresIn = json['expires_in'];  // seconds until token expires.
    const resourceTokenExpiresAt = moment().add(expiresIn, 'seconds');
    ServerStorage.persist('ResourceTokenExpiresAt', resourceTokenExpiresAt);
    return resourceToken;
    ```

1. <span data-ttu-id="0cc95-288">ファイルを閉じないで保存します。</span><span class="sxs-lookup"><span data-stu-id="0cc95-288">Save the file, but don't close it.</span></span>

### <a name="create-a-method-to-get-access-to-the-resource-using-the-on-behalf-of-flow"></a><span data-ttu-id="0cc95-289">「代理」ワークフローを使用してリソースにアクセスするメソッドを作成する</span><span class="sxs-lookup"><span data-stu-id="0cc95-289">Create a method to get access to the resource using the "on behalf of" flow</span></span>

1. <span data-ttu-id="0cc95-p155">引き続き src/auth.ts で、次に示すメソッドを `AuthModule` クラスに追加します。このコードについては、以下に注意してください。</span><span class="sxs-lookup"><span data-stu-id="0cc95-p155">Still in src/auth.ts, add the method below to the `AuthModule` class. Note the following about this code:</span></span>

    * <span data-ttu-id="0cc95-292">`exchangeForToken` メソッドへのパラメーターに関する上記のコメントは、このメソッドのパラメーターにも当てはまります。</span><span class="sxs-lookup"><span data-stu-id="0cc95-292">The comments above about the parameters to the the `exchangeForToken` method apply to the parameters of this method as well.</span></span>
    * <span data-ttu-id="0cc95-p156">このメソッドでは、最初にリソースへの有効期限が切れていない (次の 1 分まで有効期限が続く) アクセス トークンについて永続ストレージをチェックします。これは、直前のセクションで作成した `exchangeForToken` メソッドを呼び出します (そのメソッドが必要になる場合)。</span><span class="sxs-lookup"><span data-stu-id="0cc95-p156">The method first checks the persistent storage for an access token to the resource that has not expired and is not going to expire in the next minute. It calls the `exchangeForToken` method you created in the last section only if it needs to.</span></span>

    ```typescript
    async acquireTokenOnBehalfOf(jwt: string, scopes: string[] = ['openid'], resource?: string) {
        const resourceTokenExpirationTime = ServerStorage.retrieve('ResourceTokenExpiresAt');
        if (moment().add(1, 'minute').diff(await resourceTokenExpirationTime) < 1 ) {
            return ServerStorage.retrieve('ResourceToken');
        } else if (resource) {
            return this.exchangeForToken(jwt, scopes, resource);
        } else {
            return this.exchangeForToken(jwt, scopes);
        }
    }
    ```

1. <span data-ttu-id="0cc95-295">ファイルを保存して閉じます。</span><span class="sxs-lookup"><span data-stu-id="0cc95-295">Save and close the file.</span></span>

### <a name="create-the-endpoints-that-will-serve-the-add-ins-home-page-and-data"></a><span data-ttu-id="0cc95-296">アドインのホーム ページとデータを提供するエンドポイントを作成する</span><span class="sxs-lookup"><span data-stu-id="0cc95-296">Create the endpoints that will serve the add-in's home page and data</span></span>

1. <span data-ttu-id="0cc95-297">src\server.ts ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="0cc95-297">Open the src\server.ts file.</span></span>

1. <span data-ttu-id="0cc95-p157">次に示すメソッドをファイルの末尾に追加します。このメソッドにより、アドインのホーム ページを提供します。アドイン マニフェストで、ホーム ページの URL を指定します。</span><span class="sxs-lookup"><span data-stu-id="0cc95-p157">Add the following method to the bottom of the file. This method will serve the add-in's home page. The add-in manifest specifies the home page URL.</span></span>

    ```typescript
    app.get('/index.html', handler(async (req, res) => {
        return res.sendfile('index.html');
    }));
    ```

1. <span data-ttu-id="0cc95-p158">ファイルの末尾に次のメソッドを追加します。このメソッドが、`values` API に対するすべての要求を処理します。</span><span class="sxs-lookup"><span data-stu-id="0cc95-p158">Add the following method to bottom of the file. This method will handle any requests for the `values` API.</span></span>

    ```typescript
    app.get('/api/values', handler(async (req, res) => {
        // TODO7: Initialize the AuthModule object and validate the access token
        //        that the client-side received from the Office host.
        // TODO8: Get a token to Microsoft Graph from either persistent storage
        //        or the "on behalf of" flow.
        // TODO9: Use the token to get data from Microsoft Graph.
        // TODO10: Relay any errors from Microsoft Graph to the client.
        // TODO11: Send to the client only the data that it actually needs.
    }));
    ```

1. <span data-ttu-id="0cc95-303">`TODO7` を次に示すコードに置き換えます。このコードは、Office ホスト アプリケーションから受け取ったアクセス トークンを検証します。</span><span class="sxs-lookup"><span data-stu-id="0cc95-303">Replace `TODO7` with the following code which validates the access token received from the Office host application.</span></span> <span data-ttu-id="0cc95-304">`verifyJWT` メソッドは、src\auth.ts ファイルで定義されています。</span><span class="sxs-lookup"><span data-stu-id="0cc95-304">The `verifyJWT` method is defined in the src\auth.ts file.</span></span> <span data-ttu-id="0cc95-305">このメソッドは、常に対象ユーザーと発行者を検証します。</span><span class="sxs-lookup"><span data-stu-id="0cc95-305">It always validates the audience and the issuer.</span></span> <span data-ttu-id="0cc95-306">省略可能なパラメーターを使用して、アクセス トークンのスコープが `access_as_user` であることを検証する必要もあるということを指定します。</span><span class="sxs-lookup"><span data-stu-id="0cc95-306">We use the optional parameter to specify that we also want it to verify that the scope in the access token is `access_as_user`.</span></span> <span data-ttu-id="0cc95-307">これは、「代理」フローによって Microsoft Graph へのアクセス トークンを取得するために、ユーザーと Office ホストが必要とする、アドインに対する唯一のアクセス許可です。</span><span class="sxs-lookup"><span data-stu-id="0cc95-307">This is the only permission to the add-in that the user and the Office host need in order to get an access token to Microsoft Graph by means of the "on behalf" flow.</span></span>

    ```typescript
    await auth.initialize();
    const { jwt } = auth.verifyJWT(req, { scp: 'access_as_user' });
    ```

    > [!NOTE]
    > <span data-ttu-id="0cc95-p160">`access_as_user` スコープのみを使用して、Office アドインの代理 (on-behalf-of) フローを処理する API を承認する必要があります。サービス内の他の API は、独自のスコープ要件が必要です。これにより、Office が取得するトークンでアクセスできるものが制限されます。</span><span class="sxs-lookup"><span data-stu-id="0cc95-p160">You should only use the `access_as_user` scope to authorize the API that handles the on-behalf-of flow for Office Add-ins. Other APIs in your service should have their own scope requirements. This limits what can be accessed with the tokens that Office acquires.</span></span>

1. <span data-ttu-id="0cc95-p161">`TODO8` を次のコードに置き換えます。このコードについては、次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="0cc95-p161">Replace `TODO8` with the following code. Note the following about this code:</span></span>

    * <span data-ttu-id="0cc95-312">`acquireTokenOnBehalfOf` の呼び出しには、resource パラメーターは含まれません。これは、resource プロパティをサポートしていない AAD V2.0 エンドポイントで `AuthModule` オブジェクト (`auth`) を作成したためです。</span><span class="sxs-lookup"><span data-stu-id="0cc95-312">The call to `acquireTokenOnBehalfOf` does not include a resource parameter because we constructed the `AuthModule` object (`auth`) with the AAD V2.0 endpoint which does not support a resource property.</span></span>
    * <span data-ttu-id="0cc95-p162">この呼び出しの 2 番目のパラメーターでは、OneDrive 上のユーザーのファイルとフォルダーのリストを取得するために、アドインが必要とするアクセス許可を指定します。 (`profile` アクセス許可は要求されません。これは、このアクセス許可が、Microsoft Graph へのアクセス トークン用のトークンでやり取りしているときではなく、Office ホストがアドインへのアクセス トークンを取得するときにだけ必要であるためです。)</span><span class="sxs-lookup"><span data-stu-id="0cc95-p162">The second parameter of the call specifies the permissions the add-in will need to get a list of the user's files and folders on OneDrive. (The `profile` permission is not requested because it is only needed when the Office host gets the access token to your add-in, not when you are trading in that token for an access token to Microsoft Graph.)</span></span>

    ```typescript
    const graphToken = await auth.acquireTokenOnBehalfOf(jwt, ['Files.Read.All']);
    ```

1. <span data-ttu-id="0cc95-p163">`TODO9` を次のコードに置き換えます。このコードについては、次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="0cc95-p163">Replace `TODO9` with the following line. Note the following about this code:</span></span>

    * <span data-ttu-id="0cc95-317">MSGraphHelper クラスは、src\msgraph-helper.ts で定義されています。</span><span class="sxs-lookup"><span data-stu-id="0cc95-317">The MSGraphHelper class is defined in src\msgraph-helper.ts.</span></span>
    * <span data-ttu-id="0cc95-318">返す必要があるデータが最小になるように、name プロパティと最初の 3 つのアイテムのみが必要なことを指定しています。</span><span class="sxs-lookup"><span data-stu-id="0cc95-318">We minimize the data that must be returned by specifying that we only want the name property and only the first 3 items.</span></span>

    ```typescript
    const graphData = await MSGraphHelper.getGraphData(graphToken, "/me/drive/root/children", "?$select=name&$top=3");
    ```

1. <span data-ttu-id="0cc95-p164">`TODO10` を次のコードに置き換えます。 このコードでは、Microsoft Graph からの "401 Unauthorized" エラーを処理します。このエラーは、期限切れのトークンまたは無効なトークンを表している可能性があります。 この事態は、トークンの永続化ロジックによって防止されているため、発生する可能性はほとんどありません  (前述のセクション「**「代理 (on-behalf-of) 」ワークフローを使用してリソースにアクセスするメソッドを作成する**」を参照してください)。この事態が発生した場合、このコードではエラー名に "Microsoft Graph error" を使用してクライアントにエラーを中継します  (前述の手順で program.js ファイルに作成した `handleClientSideErrors` メソッドを参照してください)。この後手順で ODataHelper.js ファイルに追加するコードは、Microsoft Graph からのエラーの処理に役立ちます。</span><span class="sxs-lookup"><span data-stu-id="0cc95-p164">Replace `TODO10` with the following code. Note that this code handles '401 Unauthorized" errors from Microsoft Graph which would indicate an expired or invalid token. It is very unlikely that this would ever happen since the token persisting logic should prevent it. (See the section **Create a method to get access to the resource using the "on behalf of" flow** above.) If it does happen, this code will relay the error to the client with "Microsoft Graph error" in the error name. (See the `handleClientSideErrors` method that you created in the program.js file in an earlier step.) Code that you add to the ODataHelper.js file in a later step helps process errors from Microsoft Graph.</span></span>

    ```typescript
    if (graphData.code) {
        if (graphData.code === 401) {
            throw new UnauthorizedError('Microsoft Graph error', graphData);
        }
    }
    ```


1. <span data-ttu-id="0cc95-p165">`TODO11` を次に示すコードに置き換えます。Microsoft Graph は、`name` プロパティのみを要求した場合でも、アイテムごとに、いくつかの OData メタデータと 1 つの **eTag** プロパティを返す点に注意してください。このコードでは、アイテムの名前のみをクライアントに送信します。</span><span class="sxs-lookup"><span data-stu-id="0cc95-p165">Replace `TODO11` with the following code. Note that Microsoft Graph returns some OData metadata and an **eTag** property for every item, even if `name` is the only property requested. The code sends only the item names to the client.</span></span>

    ```typescript
    const itemNames: string[] = [];
    const oneDriveItems: string[] = graphData['value'];
    for (let item of oneDriveItems){
        itemNames.push(item['name']);
    }
    return res.json(itemNames);
    ```

1. <span data-ttu-id="0cc95-327">ファイルを保存して閉じます。</span><span class="sxs-lookup"><span data-stu-id="0cc95-327">Save and close the file.</span></span>

### <a name="add-response-handling-to-the-odatahelper"></a><span data-ttu-id="0cc95-328">ODataHelper に応答の処理を追加する</span><span class="sxs-lookup"><span data-stu-id="0cc95-328">Add response handling to the ODataHelper</span></span>

1. <span data-ttu-id="0cc95-p166">ファイル src\odata-helper.ts を開きます。 このファイルは、ほとんど完成しています。 要求の「終了」イベントを処理するコールバックの本文が欠落しています。 `TODO` を次のコードに置き換えます。 このコードの注意点は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="0cc95-p166">Open the file src\odata-helper.ts. The file is almost complete. What's missing is the body of the callback to the handler for the request "end" event. Replace the `TODO` with the following code. About this code note:</span></span>

    * <span data-ttu-id="0cc95-p167">OData エンドポイントからの応答は、エラーである可能性があります。たとえば、エンドポイントがアクセス トークンを必要としていて、そのトークンが無効または有効期限切れの場合は 401 になります。 ただし、エラー メッセージは \*\* の呼び出しでのエラーではなく`https.get`であるため、`on('error', reject)` の最後の行 `https.get` はトリガーされません。 そのため、コードでは、成功 (200) とエラー メッセージを区別して、要求された OData またはエラー情報のどちらかを含む JSON オブジェクトを呼び出し元に送信します。</span><span class="sxs-lookup"><span data-stu-id="0cc95-p167">The response from the OData endpoint might be an error, say a 401 if the endpoint requires an access token and it was invalid or expired. But an error message is still a *message*, not an error in the call of `https.get`, so the `on('error', reject)` line at the end of `https.get` isn't triggered. So, the code distinguishes success (200) messages from error messages and sends a JSON object to the caller with either the requested OData or error information.</span></span>

    ```typescript
    var error;
    if (response.statusCode === 200) {
        // TODO1: Return the data to the caller and resolve the Promise.
    } else {
       // TODO2: Return an error object to the caller and resolve the Promise.
    }
    ```

1. <span data-ttu-id="0cc95-p168">`TODO1` を次のコードと置き換えます。このコードでは、データが JSON として返されることを前提としています。</span><span class="sxs-lookup"><span data-stu-id="0cc95-p168">Replace `TODO1` with the following code. Note that the code assumes the data is returned as JSON.</span></span>

    ```typescript
    let parsedBody = JSON.parse(body);
    resolve(parsedBody);
    ```

1. <span data-ttu-id="0cc95-p169">`TODO2` を次のコードに置き換えます。このコードの注意点は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="0cc95-p169">Replace `TODO2` with the following code. Note about this code:</span></span>

    * <span data-ttu-id="0cc95-p170">OData ソースからのエラー応答には、常に statusCode が含まれています。また、通常は statusMessage が含まれています。 また、一部の OData ソースは、詳細な情報 (内部のコードやメッセージ、より具体的なコードやメッセージなど) を含む error プロパティも本文に追加します。</span><span class="sxs-lookup"><span data-stu-id="0cc95-p170">An error response from an OData source will always have a statusCode and usually a statusMessage. Some OData sources also add an error property to the body with further information, such as an inner, or more specific, code and message.</span></span>
    * <span data-ttu-id="0cc95-p171">Promise オブジェクトは解決されます。拒否されません。 `https.get` は、Web サービスがサーバー間の OData エンドポイントを呼び出すときに実行されます。 ただし、その呼び出しは、クライアントから Web サービスの Web API への呼び出しのコンテキストで行われます。 クライアントから Web サービスへの「外部」の要求は、「内部」の要求が拒否されると完了できなくなります。 さらに、`Error` の呼び出し元が OData エンドポイントからクライアントにエラーを中継する必要がある場合は、カスタムの `http.get` オブジェクトを含む要求も解決する必要があります。</span><span class="sxs-lookup"><span data-stu-id="0cc95-p171">The Promise object is resolved, not rejected. The `https.get` runs when a web service calls an OData endpoint server-to-server. But that call comes in the context of a call from a client to a web API in the web service. The "outer" request from the client to the web service never completes if this "inner" request is rejected. Also, resolving the request with the custom `Error` object is required if the caller of `http.get` needs to relay errors from the OData endpoint to the client.</span></span>

    ```typescript
    error = new Error();
    error.code = response.statusCode;
    error.message = response.statusMessage;

    // The error body sometimes includes an empty space
    // before the first character, remove it or it causes an error.
    body = body.trim();
    error.bodyCode = JSON.parse(body).error.code;
    error.bodyMessage = JSON.parse(body).error.message;
    resolve(error);
    ```

1. <span data-ttu-id="0cc95-348">ファイルを保存して閉じます。</span><span class="sxs-lookup"><span data-stu-id="0cc95-348">Save and close the file.</span></span>

## <a name="deploy-the-add-in"></a><span data-ttu-id="0cc95-349">アドインを展開する</span><span class="sxs-lookup"><span data-stu-id="0cc95-349">Deploy the add-in</span></span>

<span data-ttu-id="0cc95-350">次に、Office がアドインを検索する場所を認識できるようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="0cc95-350">Now you need to let Office know where to find the add-in.</span></span>

1. <span data-ttu-id="0cc95-351">ネットワーク共有を作成するか、[フォルダーをネットワークに共有します](/previous-versions/windows/it-pro/windows-server-2008-R2-and-2008/cc770880(v=ws.11))。</span><span class="sxs-lookup"><span data-stu-id="0cc95-351">Create a network share, or [share a folder to the network](/previous-versions/windows/it-pro/windows-server-2008-R2-and-2008/cc770880(v=ws.11)).</span></span>

1. <span data-ttu-id="0cc95-352">プロジェクトのルートから、Office-Add-in-NodeJS-SSO.xml マニフェスト ファイルのコピーを共有フォルダーに配置します。</span><span class="sxs-lookup"><span data-stu-id="0cc95-352">Place a copy of the Office-Add-in-NodeJS-SSO.xml manifest file, from the root of the project, into the shared folder.</span></span>

1. <span data-ttu-id="0cc95-353">PowerPoint を起動して、ドキュメントを開きます。</span><span class="sxs-lookup"><span data-stu-id="0cc95-353">Launch PowerPoint and open a document.</span></span>

1. <span data-ttu-id="0cc95-354">**[ファイル]** タブを選択して、**[オプション]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="0cc95-354">Choose the **File** tab, and then choose **Options**.</span></span>

1. <span data-ttu-id="0cc95-355">[**セキュリティ センター**] を選択し、[**セキュリティ センターの設定**] ボタンを選択します。</span><span class="sxs-lookup"><span data-stu-id="0cc95-355">Choose **Trust Center**, and then choose the **Trust Center Settings** button.</span></span>

1. <span data-ttu-id="0cc95-356">**[信頼されているアドイン カタログ]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="0cc95-356">Choose **Trusted Add-ins Catalogs**.</span></span>

1. <span data-ttu-id="0cc95-357">**[カタログの URL]** フィールドに、Office-Add-in-NodeJS-SSO.xml があるフォルダー共有へのネットワーク パスを入力して、**[カタログの追加]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="0cc95-357">In the **Catalog Url** field, enter the network path to the folder share that contains Office-Add-in-NodeJS-SSO.xml, and then choose **Add Catalog**.</span></span>

1. <span data-ttu-id="0cc95-358">**[メニューに表示する]** チェック ボックスをオンにして、**[OK]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="0cc95-358">Select the **Show in Menu** check box, and then choose **OK**.</span></span>

1. <span data-ttu-id="0cc95-p172">これらの設定は Microsoft Office を次回起動したときに適用されることを示すメッセージが表示されます。PowerPoint を終了します。</span><span class="sxs-lookup"><span data-stu-id="0cc95-p172">A message is displayed to inform you that your settings will be applied the next time you start Microsoft Office. Close PowerPoint.</span></span>

## <a name="build-and-run-the-project"></a><span data-ttu-id="0cc95-361">プロジェクトのビルドと実行</span><span class="sxs-lookup"><span data-stu-id="0cc95-361">Build and run the project</span></span>

<span data-ttu-id="0cc95-p173">プロジェクトのビルドと実行には 2 つの方法があり、Visual Studio Code を使用しているかどうかによって決まります。どちらの方法でも、プロジェクトをビルドして、コードに変更を加えたときには自動的に再ビルドしてから再実行します。</span><span class="sxs-lookup"><span data-stu-id="0cc95-p173">There are two ways to build and run the project depending on whether you are using Visual Studio Code. For both ways, the project builds and automatically rebuilds and reruns when you make changes to the code.</span></span>

1. <span data-ttu-id="0cc95-364">Visual Studio Code を使用していない場合:</span><span class="sxs-lookup"><span data-stu-id="0cc95-364">If you are not using Visual Studio Code:</span></span>
   1. <span data-ttu-id="0cc95-365">ノード ターミナルを開いて、プロジェクトのルート フォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="0cc95-365">Open a node terminal and navigate to the root folder of the project.</span></span>
   1. <span data-ttu-id="0cc95-366">ターミナルで、「**npm run build**」と入力します。</span><span class="sxs-lookup"><span data-stu-id="0cc95-366">In the terminal, enter **npm run build**.</span></span>
   1. <span data-ttu-id="0cc95-367">2 番目のノード ターミナルを開いて、プロジェクトのルート フォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="0cc95-367">Open a second node terminal and navigate to the root folder of the project.</span></span>
   1. <span data-ttu-id="0cc95-368">ターミナルで、「**npm run start**」と入力します。</span><span class="sxs-lookup"><span data-stu-id="0cc95-368">In the terminal, enter **npm run start**.</span></span>

1. <span data-ttu-id="0cc95-369">VS Code を使用している場合:</span><span class="sxs-lookup"><span data-stu-id="0cc95-369">If you are using VS Code:</span></span>
   1. <span data-ttu-id="0cc95-370">VS Code でプロジェクトを開きます。</span><span class="sxs-lookup"><span data-stu-id="0cc95-370">Open the project in VS Code.</span></span>
   1. <span data-ttu-id="0cc95-371">CTRL + SHIFT + B を押して、プロジェクトをビルドします。</span><span class="sxs-lookup"><span data-stu-id="0cc95-371">Press CTRL-SHIFT-B to build the project.</span></span>
   1. <span data-ttu-id="0cc95-372">**F5** を押して、デバッグ セッションでプロジェクトを実行します。</span><span class="sxs-lookup"><span data-stu-id="0cc95-372">Press **F5** to run the project in a debugging session.</span></span>


## <a name="add-the-add-in-to-an-office-document"></a><span data-ttu-id="0cc95-373">Office ドキュメントにアドインを追加する</span><span class="sxs-lookup"><span data-stu-id="0cc95-373">Add the add-in to an Office document</span></span>

1. <span data-ttu-id="0cc95-374">PowerPoint を再起動して、プレゼンテーションを開くか作成します。</span><span class="sxs-lookup"><span data-stu-id="0cc95-374">Restart PowerPoint and open or create a presentation.</span></span>

1. <span data-ttu-id="0cc95-375">**[開発]** タブがリボンに表示されていない場合、次の手順で有効にします。</span><span class="sxs-lookup"><span data-stu-id="0cc95-375">If the **Developer** tab is not visible on the ribbon, enable it with the following steps:</span></span>
   1. <span data-ttu-id="0cc95-376">**[ファイル]**、**[オプション]**、**[リボンのユーザー設定]** の順に移動します。</span><span class="sxs-lookup"><span data-stu-id="0cc95-376">Navigate to **File** | **Options** | **Customize Ribbon**.</span></span>
   1. <span data-ttu-id="0cc95-377">チェック ボックスをオンにし、**[リボンのユーザー設定]** ページの右にあるコントロール名のツリーで **[開発]** を有効にします。</span><span class="sxs-lookup"><span data-stu-id="0cc95-377">Click the check box to enable **Developer** in the tree of control names on the right of the **Customize Ribbon** page.</span></span>
   1. <span data-ttu-id="0cc95-378">**[OK]** を押します。</span><span class="sxs-lookup"><span data-stu-id="0cc95-378">Press **OK**.</span></span>

1. <span data-ttu-id="0cc95-379">PowerPoint の **[開発]** タブで、**[個人用アドイン]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="0cc95-379">On the **Developer** tab in PowerPoint, choose **My Add-ins**.</span></span>

1. <span data-ttu-id="0cc95-380">**[共有フォルダー]** タブを選択します。</span><span class="sxs-lookup"><span data-stu-id="0cc95-380">Select the **SHARED FOLDER** tab.</span></span>

1. <span data-ttu-id="0cc95-381">**[SSO NodeJS Sample]** を選択して、**[OK]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="0cc95-381">Choose **SSO NodeJS Sample**, and then select **OK**.</span></span>

1. <span data-ttu-id="0cc95-382">**[ホーム]** リボンに、**[SSO NodeJS]** という新しいグループが表示され、**[アドインの表示]** というラベルの付いたボタンとアイコンが含まれています。</span><span class="sxs-lookup"><span data-stu-id="0cc95-382">On the **Home** ribbon is a new group called **SSO NodeJS** with a button labeled **Show Add-in** and an icon.</span></span>

## <a name="test-the-add-in"></a><span data-ttu-id="0cc95-383">アドインをテストする</span><span class="sxs-lookup"><span data-stu-id="0cc95-383">Test the add-in</span></span>

1. <span data-ttu-id="0cc95-384">結果を確認できるように、OneDrive 内にファイルがいくつかあることを確認します。</span><span class="sxs-lookup"><span data-stu-id="0cc95-384">Ensure that you have some files in your OneDrive so that you can verify the results.</span></span>

1. <span data-ttu-id="0cc95-385">**[アドインの表示]** ボタンをクリックして、アドインを開きます。</span><span class="sxs-lookup"><span data-stu-id="0cc95-385">Click **Show Add-in** button to open the add-in.</span></span>

1. <span data-ttu-id="0cc95-p174">[ようこそ] ページでアドインが開きます。**[OneDrive からファイルを取得]** ボタンをクリックします。</span><span class="sxs-lookup"><span data-stu-id="0cc95-p174">The add-in opens with a Welcome page. Click the **Get My Files from OneDrive** button.</span></span>

1. <span data-ttu-id="0cc95-p175">Office にサインインしている場合は、このボタンの下に OneDrive にあるファイルとフォルダーのリストが表示されます。これは、初回実行時には 15 秒以上かかることがあります。</span><span class="sxs-lookup"><span data-stu-id="0cc95-p175">If you are are signed into Office, a list of your files and folders on OneDrive will appear below the button. This may take more than 15 seconds the first time.</span></span>

1. <span data-ttu-id="0cc95-390">Office にサインインしていない場合は、ポップアップが表示され、サインインするように求められます。</span><span class="sxs-lookup"><span data-stu-id="0cc95-390">If you are not signed into Office, a popup will open and prompt you to sign in.</span></span> <span data-ttu-id="0cc95-391">サインインが完了すると、数秒後にファイルとフォルダーの一覧が表示されます。</span><span class="sxs-lookup"><span data-stu-id="0cc95-391">After you have completed the sign-in, the list of your files and folders will appear after a few seconds.</span></span> <span data-ttu-id="0cc95-392">*2 回目はボタンを押す必要はありません。*</span><span class="sxs-lookup"><span data-stu-id="0cc95-392">*You should not press the button a second time.*</span></span>

> [!NOTE]
> <span data-ttu-id="0cc95-p177">以前に別の ID で Office にサインオンしていて、そのときに開いたいくつかの Office アプリケーションが引き続き開いている場合、Office がその ID を確実に変更するとは限りません (PowerPoint で ID が変更済みのように表示されている場合でも)。 このような場合は、Microsoft Graph への呼び出しが失敗するか、以前の ID からのデータが返される可能性があります。 これを防止するには、必ず*他のすべての Office アプリケーションを閉じて*から、**[OneDrive からファイルを取得]** を押します。</span><span class="sxs-lookup"><span data-stu-id="0cc95-p177">If you were previously signed on to Office with a different ID, and some Office applications that were open at the time are still open, Office may not reliably change your ID even if it appears to have done so in PowerPoint. If this happens, the call to Microsoft Graph may fail or data from the previous ID may be returned. To prevent this, be sure to *close all other Office applications* before you press **Get My Files from OneDrive**.</span></span>
