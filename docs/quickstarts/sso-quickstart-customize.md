---
title: Node.js SSO が有効なアドインをカスタマイズする
description: '[ごみ箱] ジェネレーターを使用して作成した SSO が有効なアドインのカスタマイズについて説明します。'
ms.date: 07/07/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: c1d292ed8ead40201dd035d6ae8e6997174ea477
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094485"
---
# <a name="customize-your-nodejs-sso-enabled-add-in"></a><span data-ttu-id="043a3-103">Node.js SSO が有効なアドインをカスタマイズする</span><span class="sxs-lookup"><span data-stu-id="043a3-103">Customize your Node.js SSO-enabled add-in</span></span>

> [!IMPORTANT]
> <span data-ttu-id="043a3-104">この記事は、[シングルサインオン (sso) のクイックスタート](sso-quickstart.md)を完了して作成された sso が有効なアドインに基づいて構築されています。</span><span class="sxs-lookup"><span data-stu-id="043a3-104">This article builds upon the SSO-enabled add-in that's created by completing the [single sign-on (SSO) quick start](sso-quickstart.md).</span></span> <span data-ttu-id="043a3-105">この記事を読む前に、クイックスタートを完了してください。</span><span class="sxs-lookup"><span data-stu-id="043a3-105">Please complete the quick start before reading this article.</span></span>

<span data-ttu-id="043a3-106">[Sso クイックスタート](sso-quickstart.md)では、サインインしているユーザーのプロファイル情報を取得し、それをドキュメントまたはメッセージに書き込む sso が有効なアドインを作成します。</span><span class="sxs-lookup"><span data-stu-id="043a3-106">The [SSO quick start](sso-quickstart.md) creates an SSO-enabled add-in that gets the signed-in user's profile information and writes it to the document or message.</span></span> <span data-ttu-id="043a3-107">この記事では、SSO クイックスタートで、[ごみ箱] ジェネレーターを使用して作成したアドインを更新するプロセスについて説明し、別のアクセス許可を必要とする新しい機能を追加します。</span><span class="sxs-lookup"><span data-stu-id="043a3-107">In this article, you'll walk through the process of updating the add-in that you created with the Yeoman generator in the SSO quick start, to add new functionality that requires different permissions.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="043a3-108">前提条件</span><span class="sxs-lookup"><span data-stu-id="043a3-108">Prerequisites</span></span>

* <span data-ttu-id="043a3-109">[SSO クイックスタート](sso-quickstart.md)の手順に従って作成した Office アドイン。</span><span class="sxs-lookup"><span data-stu-id="043a3-109">An Office Add-in that you created by following the instructions in the [SSO quick start](sso-quickstart.md).</span></span>

* <span data-ttu-id="043a3-110">少なくとも、Microsoft 365 サブスクリプションの OneDrive for Business に格納されているファイルとフォルダーがいくつかあります。</span><span class="sxs-lookup"><span data-stu-id="043a3-110">At least a few files and folders stored on OneDrive for Business in your Microsoft 365 subscription.</span></span>

* <span data-ttu-id="043a3-111">[Node.js](https://nodejs.org) (最新 [LTS](https://nodejs.org/about/releases) バージョン)。</span><span class="sxs-lookup"><span data-stu-id="043a3-111">[Node.js](https://nodejs.org) (the latest [LTS](https://nodejs.org/about/releases) version).</span></span>

[!include[additional prerequisites](../includes/sso-tutorial-prereqs.md)]

## <a name="review-contents-of-the-project"></a><span data-ttu-id="043a3-112">プロジェクトのコンテンツをレビューする</span><span class="sxs-lookup"><span data-stu-id="043a3-112">Review contents of the project</span></span>

<span data-ttu-id="043a3-113">まず、以前に[使用](sso-quickstart.md)していたアドインプロジェクトのクイックレビューから始めましょう。</span><span class="sxs-lookup"><span data-stu-id="043a3-113">Let's begin with a quick review of the add-in project that you previously [created with the Yeoman generator](sso-quickstart.md).</span></span>

> [!NOTE]
> <span data-ttu-id="043a3-114">この記事では、ファイル拡張子 **.js**を使用してスクリプトファイルを参照する場所で、プロジェクトが TypeScript を使用して作成されている場合は **、ファイル拡張子**としてを指定します。</span><span class="sxs-lookup"><span data-stu-id="043a3-114">In places where this article references script files using **.js** file extension, assume the **.ts** file extension instead if your project was created with TypeScript.</span></span>

[!include[project structure for an SSO-enabled add-in created with the Yeoman generator](../includes/sso-yeoman-project-structure.md)]

## <a name="add-new-functionality"></a><span data-ttu-id="043a3-115">新しい機能を追加する</span><span class="sxs-lookup"><span data-stu-id="043a3-115">Add new functionality</span></span>

<span data-ttu-id="043a3-116">SSO クイックスタートを使用して作成したアドインは、Microsoft Graph を使用してサインインしているユーザーのプロファイル情報を取得し、その情報をドキュメントまたはメッセージに書き込みます。</span><span class="sxs-lookup"><span data-stu-id="043a3-116">The add-in that you created with the SSO quick start uses Microsoft Graph to get the signed-in user's profile information and writes that information to the document or message.</span></span> <span data-ttu-id="043a3-117">サインインしているユーザーの OneDrive for Business から上位10個のファイルとフォルダーの名前を取得し、その情報をドキュメントまたはメッセージに書き込むようにアドインの機能を変更しましょう。</span><span class="sxs-lookup"><span data-stu-id="043a3-117">Let's change the add-in's functionality such that it gets the names of the top 10 files and folders from the signed-in user's OneDrive for Business and writes that information to the document or message.</span></span> <span data-ttu-id="043a3-118">この新しい機能を有効にするには、Azure でアプリのアクセス許可を更新する必要があります。また、アドインプロジェクト内のコードを更新する必要があります。</span><span class="sxs-lookup"><span data-stu-id="043a3-118">Enabling this new functionality requires updating app permissions in Azure and updating code within the add-in project.</span></span>

### <a name="update-app-permissions-in-azure"></a><span data-ttu-id="043a3-119">Azure でアプリのアクセス許可を更新する</span><span class="sxs-lookup"><span data-stu-id="043a3-119">Update app permissions in Azure</span></span>

<span data-ttu-id="043a3-120">アドインがユーザーの OneDrive for Business のコンテンツを正常に読み取る前に、Azure のアプリ登録情報を適切なアクセス許可で更新する必要があります。</span><span class="sxs-lookup"><span data-stu-id="043a3-120">Before the add-in can successfully read the contents of the user's OneDrive for Business, its app registration information in Azure must be updated with the appropriate permissions.</span></span> <span data-ttu-id="043a3-121">次の手順を実行して、アプリに**ファイルの読み取り**アクセス許可を付与し、ユーザーを取り消し**ます。読み取り**アクセス許可は不要になりました。</span><span class="sxs-lookup"><span data-stu-id="043a3-121">Complete the following steps to grant the app the **Files.Read.All** permission and revoke the **User.Read** permission, which is no longer needed.</span></span>

1. <span data-ttu-id="043a3-122">[Azure portal](https://ms.portal.azure.com/#home)に移動し、 **Microsoft 365 管理者の資格情報を使用してサインイン**します。</span><span class="sxs-lookup"><span data-stu-id="043a3-122">Navigate to the [Azure portal](https://ms.portal.azure.com/#home) and **sign in using your Microsoft 365 administrator credentials**.</span></span>

2. <span data-ttu-id="043a3-123">[アプリの**登録**] ページに移動します。</span><span class="sxs-lookup"><span data-stu-id="043a3-123">Navigate to the **App registrations** page.</span></span>
    > [!TIP]
    > <span data-ttu-id="043a3-124">これを行うには、Azure ホームページで**アプリ登録**タイルを選択するか、ホームページの検索ボックスを使用して**アプリの登録**を見つけて選択します。</span><span class="sxs-lookup"><span data-stu-id="043a3-124">You can do this either by choosing the **App registrations** tile on the Azure home page or by using the search box on the home page to find and choose **App registrations**.</span></span>

3. <span data-ttu-id="043a3-125">[**アプリの登録**] ページで、クイックスタート時に作成したアプリを選択します。</span><span class="sxs-lookup"><span data-stu-id="043a3-125">On the **App registrations** page, choose the app that you created during the quick start.</span></span> 
    > [!TIP]
    > <span data-ttu-id="043a3-126">アプリの**表示名**は、そのプロジェクトの作成時に指定したアドイン名と一致します。</span><span class="sxs-lookup"><span data-stu-id="043a3-126">The **Display name** of the app will match the add-in name that you specified when you created the project with the Yeoman generator.</span></span>

4. <span data-ttu-id="043a3-127">[アプリの概要] ページで、ページの左側にある [**管理**] 見出しの下にある [ **API の権限**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="043a3-127">From the app overview page, choose **API permissions** under the **Manage** heading on the left side of the page.</span></span>

5. <span data-ttu-id="043a3-128">[アクセス許可] テーブルの [ユーザー] の**読み取り**行で、省略記号を選択し、表示されるメニューから [**管理者の同意を取り消す**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="043a3-128">In the **User.Read** row of the permissions table, choose the ellipsis and then select **Revoke admin consent** from the menu that appears.</span></span>

6. <span data-ttu-id="043a3-129">表示されたプロンプトに対して [**はい、削除**] ボタンを選択します。</span><span class="sxs-lookup"><span data-stu-id="043a3-129">Select the **Yes, remove** button in response to the prompt that's displayed.</span></span>

7. <span data-ttu-id="043a3-130">[アクセス許可] テーブルの [ユーザー] の**読み取り**行で、省略記号を選択し、表示されるメニューから [**アクセス許可の削除**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="043a3-130">In the **User.Read** row of the permissions table, choose the ellipsis and then select **Remove permission** from the menu that appears.</span></span>

8. <span data-ttu-id="043a3-131">表示されたプロンプトに対して [**はい、削除**] ボタンを選択します。</span><span class="sxs-lookup"><span data-stu-id="043a3-131">Select the **Yes, remove** button in response to the prompt that's displayed.</span></span>

9. <span data-ttu-id="043a3-132">**[アクセス許可の追加]** ボタンを選択します。</span><span class="sxs-lookup"><span data-stu-id="043a3-132">Select the **Add a permission** button.</span></span>

10. <span data-ttu-id="043a3-133">表示されたパネルで、[ **Microsoft Graph** ] を選択し、[**代理アクセス許可**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="043a3-133">On the panel that opens choose **Microsoft Graph** and then choose **Delegated permissions**.</span></span>

11. <span data-ttu-id="043a3-134">[ **API アクセス許可の要求**] パネルで、次のようにします。</span><span class="sxs-lookup"><span data-stu-id="043a3-134">On the **Request API permissions** panel:</span></span>

    <span data-ttu-id="043a3-135">a. </span><span class="sxs-lookup"><span data-stu-id="043a3-135">a.</span></span> <span data-ttu-id="043a3-136">[**ファイル**] の下で、[ファイル] を選択します **。**</span><span class="sxs-lookup"><span data-stu-id="043a3-136">Under **Files**, select **Files.Read.All**.</span></span>

    <span data-ttu-id="043a3-137">b. </span><span class="sxs-lookup"><span data-stu-id="043a3-137">b.</span></span> <span data-ttu-id="043a3-138">パネルの下部にある [**アクセス許可の追加**] ボタンを選択して、これらのアクセス許可の変更を保存します。</span><span class="sxs-lookup"><span data-stu-id="043a3-138">Select the **Add permissions** button at the bottom of the panel to save these permissions changes.</span></span>

12. <span data-ttu-id="043a3-139">**[[テナント名] に対する管理者の同意を許可**する] ボタンを選択します。</span><span class="sxs-lookup"><span data-stu-id="043a3-139">Select the **Grant admin consent for [tenant name]** button.</span></span>

13. <span data-ttu-id="043a3-140">表示されるプロンプトに対して [**はい**] ボタンを選択します。</span><span class="sxs-lookup"><span data-stu-id="043a3-140">Select the **Yes** button in response to the prompt that's displayed.</span></span>

### <a name="update-code-in-the-add-in-project"></a><span data-ttu-id="043a3-141">アドインプロジェクトでコードを更新する</span><span class="sxs-lookup"><span data-stu-id="043a3-141">Update code in the add-in project</span></span>

<span data-ttu-id="043a3-142">サインインしているユーザーの OneDrive for Business の内容をアドインが読み取ることができるようにするには、次のことを行う必要があります。</span><span class="sxs-lookup"><span data-stu-id="043a3-142">To enable the add-in to read contents of the signed-in user's OneDrive for Business, you'll need to:</span></span>

- <span data-ttu-id="043a3-143">Microsoft Graph の URL、パラメーター、および必要なアクセススコープを参照するコードを更新します。</span><span class="sxs-lookup"><span data-stu-id="043a3-143">Update the code that references the Microsoft Graph URL, parameters, and required access scope.</span></span>

- <span data-ttu-id="043a3-144">作業ウィンドウの UI を定義するコードを更新して、新しい機能を正確に記述できるようにします。</span><span class="sxs-lookup"><span data-stu-id="043a3-144">Update the code that defines the task pane UI, so that it accurately describes the new functionality.</span></span> 

- <span data-ttu-id="043a3-145">Microsoft Graph から応答を解析するコードを更新し、ドキュメントまたはメッセージに書き込みます。</span><span class="sxs-lookup"><span data-stu-id="043a3-145">Update the code that parses the response from Microsoft Graph and writes it to the document or message.</span></span>

<span data-ttu-id="043a3-146">次の手順では、これらの更新について説明します。</span><span class="sxs-lookup"><span data-stu-id="043a3-146">The following steps describe these updates.</span></span>

### <a name="changes-required-for-any-type-of-add-in"></a><span data-ttu-id="043a3-147">任意の種類のアドインに必要な変更</span><span class="sxs-lookup"><span data-stu-id="043a3-147">Changes required for any type of add-in</span></span>

<span data-ttu-id="043a3-148">アドインに対して次の手順を実行して、Microsoft Graph の URL、パラメーター、およびアクセススコープを変更し、作業ウィンドウの UI を更新します。</span><span class="sxs-lookup"><span data-stu-id="043a3-148">Complete the following steps for your add-in, to change the Microsoft Graph URL, parameters, and access scope, and update the taskpane UI.</span></span> <span data-ttu-id="043a3-149">これらの手順は、アドインの対象となる Office ホストに関係なく同じです。</span><span class="sxs-lookup"><span data-stu-id="043a3-149">These steps are the same, regardless of which Office host your add-in targets.</span></span>

1. <span data-ttu-id="043a3-150">**./.ENV**ファイル:</span><span class="sxs-lookup"><span data-stu-id="043a3-150">In the **./.ENV** file:</span></span>

    <span data-ttu-id="043a3-151">a. </span><span class="sxs-lookup"><span data-stu-id="043a3-151">a.</span></span> <span data-ttu-id="043a3-152">`GRAPH_URL_SEGMENT=/me`を次のように置き換えます。`GRAPH_URL_SEGMENT=/me/drive/root/children`</span><span class="sxs-lookup"><span data-stu-id="043a3-152">Replace `GRAPH_URL_SEGMENT=/me` with the following: `GRAPH_URL_SEGMENT=/me/drive/root/children`</span></span>

    <span data-ttu-id="043a3-153">b. </span><span class="sxs-lookup"><span data-stu-id="043a3-153">b.</span></span> <span data-ttu-id="043a3-154">`QUERY_PARAM_SEGMENT=`を次のように置き換えます。`QUERY_PARAM_SEGMENT=?$select=name&$top=10`</span><span class="sxs-lookup"><span data-stu-id="043a3-154">Replace `QUERY_PARAM_SEGMENT=` with the following: `QUERY_PARAM_SEGMENT=?$select=name&$top=10`</span></span>

    <span data-ttu-id="043a3-155">c.</span><span class="sxs-lookup"><span data-stu-id="043a3-155">c.</span></span> <span data-ttu-id="043a3-156">`SCOPE=User.Read`を次のように置き換えます。`SCOPE=Files.Read.All`</span><span class="sxs-lookup"><span data-stu-id="043a3-156">Replace `SCOPE=User.Read` with the following: `SCOPE=Files.Read.All`</span></span>

2. <span data-ttu-id="043a3-157">**./manifest.xml**で、 `<Scope>User.Read</Scope>` ファイルの末尾付近の行を見つけて行に置き換え `<Scope>Files.Read.All</Scope>` ます。</span><span class="sxs-lookup"><span data-stu-id="043a3-157">In **./manifest.xml**, find the line `<Scope>User.Read</Scope>` near the end of the file and replace it with the line `<Scope>Files.Read.All</Scope>`.</span></span>

3. <span data-ttu-id="043a3-158">**/Src/helpers/fallbackauthdialog.js** (または TypeScript プロジェクトの **/src/helpers/fallbackauthdialog.ts** ) で、文字列を見つけて、次のように定義され `https://graph.microsoft.com/User.Read` た文字列で置き換え `https://graph.microsoft.com/Files.Read.All` `requestObj` ます。</span><span class="sxs-lookup"><span data-stu-id="043a3-158">In **./src/helpers/fallbackauthdialog.js** (or in **./src/helpers/fallbackauthdialog.ts** for a TypeScript project), find the string `https://graph.microsoft.com/User.Read` and replace it with the string `https://graph.microsoft.com/Files.Read.All`, such that `requestObj` is defined as follows:</span></span>

    ```javascript
    var requestObj = {
      scopes: [`https://graph.microsoft.com/Files.Read.All`]
    };
    ```

    ```typescript
    var requestObj: Object = {
      scopes: [`https://graph.microsoft.com/Files.Read.All`]
    };
    ```

4. <span data-ttu-id="043a3-159">**/Src/taskpane/taskpane.html**で、要素を検索し、その要素内のテキストを更新して、 `<section class="ms-firstrun-instructionstep__header">` アドインの新しい機能を記述します。</span><span class="sxs-lookup"><span data-stu-id="043a3-159">In **./src/taskpane/taskpane.html**, find the element `<section class="ms-firstrun-instructionstep__header">` and update the text within that element to describe the add-in's new functionality.</span></span>

    ```html
    <section class="ms-firstrun-instructionstep__header">
        <h2 class="ms-font-m">This add-in demonstrates how to use single sign-on by making a call to Microsoft
            Graph to read content from OneDrive for Business.</h2>
        <div class="ms-firstrun-instructionstep__header--image"></div>
    </section>
    ```

5. <span data-ttu-id="043a3-160">**./Src/taskpane/taskpane.html**で、文字列を検索し、文字列に置き換え `Get My User Profile Information` `Read my OneDrive for Business` ます。</span><span class="sxs-lookup"><span data-stu-id="043a3-160">In **./src/taskpane/taskpane.html**, find and replace both occurrences of the string `Get My User Profile Information` with the string `Read my OneDrive for Business`.</span></span>

    ```html
    <li class="ms-ListItem">
        <span class="ms-ListItem-primaryText">Click the <b>Read my OneDrive for Business</b>
            button.</span>
        <div class="clearfix"></div>
    </li>
    ```

    ```html
    <p align="center">
        <button id="getGraphDataButton" class="popupButton ms-Button ms-Button--primary"><span
                class="ms-Button-label">Read my OneDrive for Business</span></button>
    </p>
    ```

6. <span data-ttu-id="043a3-161">**/Src/taskpane/taskpane.html**で、文字列を検索して置換し `Your user profile information will be displayed in the document.` ます。 `The names of the top 10 files and folders in your OneDrive for Business will be displayed in the document or message.`</span><span class="sxs-lookup"><span data-stu-id="043a3-161">In **./src/taskpane/taskpane.html**, find and replace the string `Your user profile information will be displayed in the document.` with the string `The names of the top 10 files and folders in your OneDrive for Business will be displayed in the document or message.`.</span></span>

    ```html
    <li class="ms-ListItem">
        <span class="ms-ListItem-primaryText">The names of the top 10 files and folders in your OneDrive for Business will be displayed in the document or message.</span>
        <div class="clearfix"></div>
    </li>
    ```

7. <span data-ttu-id="043a3-162">アドインの種類に対応するセクションのガイダンスに従って、Microsoft Graph から応答を解析するコードを更新し、ドキュメントまたはメッセージに書き込みます。</span><span class="sxs-lookup"><span data-stu-id="043a3-162">Update the code that parses the response from Microsoft Graph and writes it to the document or message, by following guidance in the section that corresponds to your type of add-in:</span></span>

    - [<span data-ttu-id="043a3-163">Excel アドインに必要な変更 (JavaScript)</span><span class="sxs-lookup"><span data-stu-id="043a3-163">Changes required for an Excel add-in (JavaScript)</span></span>](#changes-required-for-an-excel-add-in-javascript)
    - [<span data-ttu-id="043a3-164">Excel アドイン (TypeScript) に必要な変更</span><span class="sxs-lookup"><span data-stu-id="043a3-164">Changes required for an Excel add-in (TypeScript)</span></span>](#changes-required-for-an-excel-add-in-typescript)
    - [<span data-ttu-id="043a3-165">Outlook アドインに必要な変更 (JavaScript)</span><span class="sxs-lookup"><span data-stu-id="043a3-165">Changes required for an Outlook add-in (JavaScript)</span></span>](#changes-required-for-an-outlook-add-in-javascript)
    - [<span data-ttu-id="043a3-166">Outlook アドイン (TypeScript) に必要な変更</span><span class="sxs-lookup"><span data-stu-id="043a3-166">Changes required for an Outlook add-in (TypeScript)</span></span>](#changes-required-for-an-outlook-add-in-typescript)
    - [<span data-ttu-id="043a3-167">PowerPoint アドインに必要な変更 (JavaScript)</span><span class="sxs-lookup"><span data-stu-id="043a3-167">Changes required for a PowerPoint add-in (JavaScript)</span></span>](#changes-required-for-a-powerpoint-add-in-javascript)
    - [<span data-ttu-id="043a3-168">PowerPoint アドイン (TypeScript) に必要な変更</span><span class="sxs-lookup"><span data-stu-id="043a3-168">Changes required for a PowerPoint add-in (TypeScript)</span></span>](#changes-required-for-a-powerpoint-add-in-typescript)
    - [<span data-ttu-id="043a3-169">Word アドインに必要な変更 (JavaScript)</span><span class="sxs-lookup"><span data-stu-id="043a3-169">Changes required for a Word add-in (JavaScript)</span></span>](#changes-required-for-a-word-add-in-javascript)
    - [<span data-ttu-id="043a3-170">Word アドイン (TypeScript) に必要な変更</span><span class="sxs-lookup"><span data-stu-id="043a3-170">Changes required for a Word add-in (TypeScript)</span></span>](#changes-required-for-a-word-add-in-typescript)

### <a name="changes-required-for-an-excel-add-in-javascript"></a><span data-ttu-id="043a3-171">Excel アドインに必要な変更 (JavaScript)</span><span class="sxs-lookup"><span data-stu-id="043a3-171">Changes required for an Excel add-in (JavaScript)</span></span>

<span data-ttu-id="043a3-172">アドインが JavaScript を使用して作成された Excel アドインである場合は、 **/src/helpers/documentHelper.js**で次のように変更します。</span><span class="sxs-lookup"><span data-stu-id="043a3-172">If your add-in is an Excel add-in that was created with JavaScript, make the following changes in **./src/helpers/documentHelper.js**:</span></span>

1. <span data-ttu-id="043a3-173">関数を検索 `writeDataToOfficeDocument` し、次の関数に置き換えます。</span><span class="sxs-lookup"><span data-stu-id="043a3-173">Find the `writeDataToOfficeDocument` function and replace it with the following function:</span></span>

    ```javascript
    export function writeDataToOfficeDocument(result) {
      return new OfficeExtension.Promise(function(resolve, reject) {
        try {
          writeDataToExcel(result);
          resolve();
        } catch (error) {
          reject(Error("Unable to write data to document. " + error.toString()));
        }
      });
    }
    ```

2. <span data-ttu-id="043a3-174">関数を検索 `filterUserProfileInfo` し、次の関数に置き換えます。</span><span class="sxs-lookup"><span data-stu-id="043a3-174">Find the `filterUserProfileInfo` function and replace it with the following function:</span></span>

    ```javascript
    function filterOneDriveInfo(result) {
      let itemNames = [];
      let oneDriveItems = result['value'];
      for (let item of oneDriveItems) {
        itemNames.push(item['name']);
      }
      return itemNames;
    }
    ```

3. <span data-ttu-id="043a3-175">関数を検索 `writeDataToExcel` し、次の関数に置き換えます。</span><span class="sxs-lookup"><span data-stu-id="043a3-175">Find the `writeDataToExcel` function and replace it with the following function:</span></span>

    ```javascript
    function writeDataToExcel(result) {
      return Excel.run(function (context) {
        var sheet = context.workbook.worksheets.getActiveWorksheet();
        let data = [];
        let oneDriveInfo = filterOneDriveInfo(result);

        for (let i = 0; i < oneDriveInfo.length; i++) {
          if (oneDriveInfo[i] !== null) {
            let innerArray = [];
            innerArray.push(oneDriveInfo[i]);
            data.push(innerArray);
          }
        }

        const rangeAddress = `B5:B${5 + (data.length - 1)}`;
        const range = sheet.getRange(rangeAddress);
        range.values = data;
        range.format.autofitColumns();

        return context.sync();
      });
    }
    ```

4. <span data-ttu-id="043a3-176">関数を削除 `writeDataToOutlook` します。</span><span class="sxs-lookup"><span data-stu-id="043a3-176">Delete the `writeDataToOutlook` function.</span></span>

5. <span data-ttu-id="043a3-177">関数を削除 `writeDataToPowerPoint` します。</span><span class="sxs-lookup"><span data-stu-id="043a3-177">Delete the `writeDataToPowerPoint` function.</span></span>

6. <span data-ttu-id="043a3-178">関数を削除 `writeDataToWord` します。</span><span class="sxs-lookup"><span data-stu-id="043a3-178">Delete the `writeDataToWord` function.</span></span>

<span data-ttu-id="043a3-179">これらの変更を行った後で、この記事の「 [try a out](#try-it-out) 」セクションに進んで、更新されたアドインを試してみてください。</span><span class="sxs-lookup"><span data-stu-id="043a3-179">After you've made these changes, skip ahead to the [Try it out](#try-it-out) section of this article to try out your updated add-in.</span></span>

### <a name="changes-required-for-an-excel-add-in-typescript"></a><span data-ttu-id="043a3-180">Excel アドイン (TypeScript) に必要な変更</span><span class="sxs-lookup"><span data-stu-id="043a3-180">Changes required for an Excel add-in (TypeScript)</span></span>

<span data-ttu-id="043a3-181">アドインが TypeScript を使用して作成された Excel アドインである場合は、 **/src/taskpane/taskpane.ts**を開き、 `writeDataToOfficeDocument` 関数を見つけて、次の関数で置き換えます。</span><span class="sxs-lookup"><span data-stu-id="043a3-181">If your add-in is an Excel add-in that was created with TypeScript, open **./src/taskpane/taskpane.ts**, find the `writeDataToOfficeDocument` function, and replace it with the following function:</span></span>

```typescript
export function writeDataToOfficeDocument(result: Object): Promise<any> {
  return Excel.run(function(context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    let data: string[] = [];

    let itemNames: string[] = [];
    let oneDriveItems = result["value"];
    for (let item of oneDriveItems) {
      itemNames.push(item["name"]);
    }

    for (let i = 0; i < itemNames.length; i++) {
      if (itemNames[i] !== null) {
        let innerArray = [];
        innerArray.push(itemNames[i]);
        data.push(innerArray);
      }
    }
    
    const rangeAddress = `B5:B${5 + (data.length - 1)}`;
    const range = sheet.getRange(rangeAddress);
    range.values = data;
    range.format.autofitColumns();

    return context.sync();
  });
}
```

<span data-ttu-id="043a3-182">これらの変更を行った後で、この記事の「 [try a out](#try-it-out) 」セクションに進んで、更新されたアドインを試してみてください。</span><span class="sxs-lookup"><span data-stu-id="043a3-182">After you've made these changes, skip ahead to the [Try it out](#try-it-out) section of this article to try out your updated add-in.</span></span>

### <a name="changes-required-for-an-outlook-add-in-javascript"></a><span data-ttu-id="043a3-183">Outlook アドインに必要な変更 (JavaScript)</span><span class="sxs-lookup"><span data-stu-id="043a3-183">Changes required for an Outlook add-in (JavaScript)</span></span>

<span data-ttu-id="043a3-184">アドインが JavaScript を使用して作成された Outlook アドインの場合は、 **/src/helpers/documentHelper.js**で次のように変更します。</span><span class="sxs-lookup"><span data-stu-id="043a3-184">If your add-in is an Outlook add-in that was created with JavaScript, make the following changes in **./src/helpers/documentHelper.js**:</span></span>

1. <span data-ttu-id="043a3-185">関数を検索 `writeDataToOfficeDocument` し、次の関数に置き換えます。</span><span class="sxs-lookup"><span data-stu-id="043a3-185">Find the `writeDataToOfficeDocument` function and replace it with the following function:</span></span>

    ```javascript
    export function writeDataToOfficeDocument(result) {
      return new OfficeExtension.Promise(function(resolve, reject) {
        try {
          writeDataToOutlook(result);
          resolve();
        } catch (error) {
          reject(Error("Unable to write data to message. " + error.toString()));
        }
      });
    }
    ```

2. <span data-ttu-id="043a3-186">関数を検索 `filterUserProfileInfo` し、次の関数に置き換えます。</span><span class="sxs-lookup"><span data-stu-id="043a3-186">Find the `filterUserProfileInfo` function and replace it with the following function:</span></span>

    ```javascript
    function filterOneDriveInfo(result) {
      let itemNames = [];
      let oneDriveItems = result['value'];
      for (let item of oneDriveItems) {
        itemNames.push(item['name']);
      }
      return itemNames;
    }
    ```

3. <span data-ttu-id="043a3-187">関数を検索 `writeDataToOutlook` し、次の関数に置き換えます。</span><span class="sxs-lookup"><span data-stu-id="043a3-187">Find the `writeDataToOutlook` function and replace it with the following function:</span></span>

    ```javascript
    function writeDataToOutlook(result) {
      let data = [];
      let oneDriveInfo = filterOneDriveInfo(result);

      for (let i = 0; i < oneDriveInfo.length; i++) {
        if (oneDriveInfo[i] !== null) {
          data.push(oneDriveInfo[i]);
        }
      }

      let objectNames = "";
      for (let i = 0; i < data.length; i++) {
        objectNames += data[i] + "<br/>";
      }

      Office.context.mailbox.item.body.setSelectedDataAsync(objectNames, { coercionType: Office.CoercionType.Html });
    }
    ```

4. <span data-ttu-id="043a3-188">関数を削除 `writeDataToExcel` します。</span><span class="sxs-lookup"><span data-stu-id="043a3-188">Delete the `writeDataToExcel` function.</span></span>

5. <span data-ttu-id="043a3-189">関数を削除 `writeDataToPowerPoint` します。</span><span class="sxs-lookup"><span data-stu-id="043a3-189">Delete the `writeDataToPowerPoint` function.</span></span>

6. <span data-ttu-id="043a3-190">関数を削除 `writeDataToWord` します。</span><span class="sxs-lookup"><span data-stu-id="043a3-190">Delete the `writeDataToWord` function.</span></span>

<span data-ttu-id="043a3-191">これらの変更を行った後で、この記事の「 [try a out](#try-it-out) 」セクションに進んで、更新されたアドインを試してみてください。</span><span class="sxs-lookup"><span data-stu-id="043a3-191">After you've made these changes, skip ahead to the [Try it out](#try-it-out) section of this article to try out your updated add-in.</span></span>

### <a name="changes-required-for-an-outlook-add-in-typescript"></a><span data-ttu-id="043a3-192">Outlook アドイン (TypeScript) に必要な変更</span><span class="sxs-lookup"><span data-stu-id="043a3-192">Changes required for an Outlook add-in (TypeScript)</span></span>

<span data-ttu-id="043a3-193">アドインが TypeScript を使用して作成された Outlook アドインの場合は、 **/src/taskpane/taskpane.ts**を開き、 `writeDataToOfficeDocument` 関数を見つけて、次の関数で置き換えます。</span><span class="sxs-lookup"><span data-stu-id="043a3-193">If your add-in is an Outlook add-in that was created with TypeScript, open **./src/taskpane/taskpane.ts**, find the `writeDataToOfficeDocument` function, and replace it with the following function:</span></span>

```typescript
export function writeDataToOfficeDocument(result: Object): void {
    let data: string[] = [];

    let itemNames: string[] = [];
    let oneDriveItems = result["value"];
    for (let item of oneDriveItems) {
        itemNames.push(item["name"]);
    };

    for (let i = 0; i < itemNames.length; i++) {
        if (itemNames[i] !== null) {
        data.push(itemNames[i]);
        }
    }

    let objectNames: string = "";
    for (let i = 0; i < data.length; i++) {
        objectNames += data[i] + "<br/>";
    }
    
    Office.context.mailbox.item.body.setSelectedDataAsync(objectNames, { coercionType: Office.CoercionType.Html });
}
```

<span data-ttu-id="043a3-194">これらの変更を行った後で、この記事の「 [try a out](#try-it-out) 」セクションに進んで、更新されたアドインを試してみてください。</span><span class="sxs-lookup"><span data-stu-id="043a3-194">After you've made these changes, skip ahead to the [Try it out](#try-it-out) section of this article to try out your updated add-in.</span></span>

### <a name="changes-required-for-a-powerpoint-add-in-javascript"></a><span data-ttu-id="043a3-195">PowerPoint アドインに必要な変更 (JavaScript)</span><span class="sxs-lookup"><span data-stu-id="043a3-195">Changes required for a PowerPoint add-in (JavaScript)</span></span>

<span data-ttu-id="043a3-196">アドインが JavaScript を使用して作成された PowerPoint アドインである場合は、 **/src/helpers/documentHelper.js**で次のように変更します。</span><span class="sxs-lookup"><span data-stu-id="043a3-196">If your add-in is a PowerPoint add-in that was created with JavaScript, make the following changes in **./src/helpers/documentHelper.js**:</span></span>

1. <span data-ttu-id="043a3-197">関数を検索 `writeDataToOfficeDocument` し、次の関数に置き換えます。</span><span class="sxs-lookup"><span data-stu-id="043a3-197">Find the `writeDataToOfficeDocument` function and replace it with the following function:</span></span>

    ```javascript
    export function writeDataToOfficeDocument(result) {
      return new OfficeExtension.Promise(function(resolve, reject) {
        try {
          writeDataToPowerPoint(result);
          resolve();
        } catch (error) {
          reject(Error("Unable to write data to document. " + error.toString()));
        }
      });
    }
    ```

2. <span data-ttu-id="043a3-198">関数を検索 `filterUserProfileInfo` し、次の関数に置き換えます。</span><span class="sxs-lookup"><span data-stu-id="043a3-198">Find the `filterUserProfileInfo` function and replace it with the following function:</span></span>

    ```javascript
    function filterOneDriveInfo(result) {
      let itemNames = [];
      let oneDriveItems = result['value'];
      for (let item of oneDriveItems) {
        itemNames.push(item['name']);
      }
      return itemNames;
    }
    ```

3. <span data-ttu-id="043a3-199">関数を検索 `writeDataToPowerPoint` し、次の関数に置き換えます。</span><span class="sxs-lookup"><span data-stu-id="043a3-199">Find the `writeDataToPowerPoint` function and replace it with the following function:</span></span>

    ```javascript
    function writeDataToPowerPoint(result) {
      let data = [];
      let oneDriveInfo = filterOneDriveInfo(result);

      for (let i = 0; i < oneDriveInfo.length; i++) {
        if (oneDriveInfo[i] !== null) {
          data.push(oneDriveInfo[i]);
        }
      }

      let objectNames = "";
      for (let i = 0; i < data.length; i++) {
        objectNames += data[i] + "\n";
      }

      Office.context.document.setSelectedDataAsync(
        objectNames, 
        function(asyncResult) {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            throw asyncResult.error.message;
          }
      });
    }
    ```

4. <span data-ttu-id="043a3-200">関数を削除 `writeDataToExcel` します。</span><span class="sxs-lookup"><span data-stu-id="043a3-200">Delete the `writeDataToExcel` function.</span></span>

5. <span data-ttu-id="043a3-201">関数を削除 `writeDataToOutlook` します。</span><span class="sxs-lookup"><span data-stu-id="043a3-201">Delete the `writeDataToOutlook` function.</span></span>

6. <span data-ttu-id="043a3-202">関数を削除 `writeDataToWord` します。</span><span class="sxs-lookup"><span data-stu-id="043a3-202">Delete the `writeDataToWord` function.</span></span>

<span data-ttu-id="043a3-203">これらの変更を行った後で、この記事の「 [try a out](#try-it-out) 」セクションに進んで、更新されたアドインを試してみてください。</span><span class="sxs-lookup"><span data-stu-id="043a3-203">After you've made these changes, skip ahead to the [Try it out](#try-it-out) section of this article to try out your updated add-in.</span></span>

### <a name="changes-required-for-a-powerpoint-add-in-typescript"></a><span data-ttu-id="043a3-204">PowerPoint アドイン (TypeScript) に必要な変更</span><span class="sxs-lookup"><span data-stu-id="043a3-204">Changes required for a PowerPoint add-in (TypeScript)</span></span>

<span data-ttu-id="043a3-205">アドインが TypeScript を使用して作成された PowerPoint アドインである場合は、 **/src/taskpane/taskpane.ts**を開き、 `writeDataToOfficeDocument` 関数を見つけて、次の関数で置き換えます。</span><span class="sxs-lookup"><span data-stu-id="043a3-205">If your add-in is a PowerPoint add-in that was created with TypeScript, open **./src/taskpane/taskpane.ts**, find the `writeDataToOfficeDocument` function, and replace it with the following function:</span></span>

```typescript
export function writeDataToOfficeDocument(result: Object): void {
  let data: string[] = [];

  let itemNames: string[] = [];
  let oneDriveItems = result["value"];
  for (let item of oneDriveItems) {
    itemNames.push(item["name"]);
  };

  for (let i = 0; i < itemNames.length; i++) {
    if (itemNames[i] !== null) {
      data.push(itemNames[i]);
    }
  }

  let objectNames: string = "";
  for (let i = 0; i < data.length; i++) {
    objectNames += data[i] + "\n";
  }

  Office.context.document.setSelectedDataAsync(objectNames, function(asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      throw asyncResult.error.message;
    }
  });
}
```

<span data-ttu-id="043a3-206">これらの変更を行った後で、この記事の「 [try a out](#try-it-out) 」セクションに進んで、更新されたアドインを試してみてください。</span><span class="sxs-lookup"><span data-stu-id="043a3-206">After you've made these changes, skip ahead to the [Try it out](#try-it-out) section of this article to try out your updated add-in.</span></span>

### <a name="changes-required-for-a-word-add-in-javascript"></a><span data-ttu-id="043a3-207">Word アドインに必要な変更 (JavaScript)</span><span class="sxs-lookup"><span data-stu-id="043a3-207">Changes required for a Word add-in (JavaScript)</span></span>

<span data-ttu-id="043a3-208">アドインが JavaScript を使用して作成された Word アドインである場合は、 **/src/helpers/documentHelper.js**で次のように変更します。</span><span class="sxs-lookup"><span data-stu-id="043a3-208">If your add-in is a Word add-in that was created with JavaScript, make the following changes in **./src/helpers/documentHelper.js**:</span></span>

1. <span data-ttu-id="043a3-209">関数を検索 `writeDataToOfficeDocument` し、次の関数に置き換えます。</span><span class="sxs-lookup"><span data-stu-id="043a3-209">Find the `writeDataToOfficeDocument` function and replace it with the following function:</span></span>

    ```javascript
    export function writeDataToOfficeDocument(result) {
      return new OfficeExtension.Promise(function(resolve, reject) {
        try {
          writeDataToWord(result);
          resolve();
        } catch (error) {
          reject(Error("Unable to write data to document. " + error.toString()));
        }
      });
    }
    ```

2. <span data-ttu-id="043a3-210">関数を検索 `filterUserProfileInfo` し、次の関数に置き換えます。</span><span class="sxs-lookup"><span data-stu-id="043a3-210">Find the `filterUserProfileInfo` function and replace it with the following function:</span></span>

    ```javascript
    function filterOneDriveInfo(result) {
      let itemNames = [];
      let oneDriveItems = result['value'];
      for (let item of oneDriveItems) {
        itemNames.push(item['name']);
      }
      return itemNames;
    }
    ```

3. <span data-ttu-id="043a3-211">関数を検索 `writeDataToWord` し、次の関数に置き換えます。</span><span class="sxs-lookup"><span data-stu-id="043a3-211">Find the `writeDataToWord` function and replace it with the following function:</span></span>

    ```javascript
    function writeDataToWord(result) {
      return Word.run(function (context) {
        let data = [];
        let oneDriveInfo = filterOneDriveInfo(result);

        for (let i = 0; i < oneDriveInfo.length; i++) {
          if (oneDriveInfo[i] !== null) {
            data.push(oneDriveInfo[i]);
          }
        }

        const documentBody = context.document.body;
        for (let i = 0; i < data.length; i++) {
          if (data[i] !== null) {
            documentBody.insertParagraph(data[i], "End");
          }
        }

        return context.sync();
      });
    }
    ```

4. <span data-ttu-id="043a3-212">関数を削除 `writeDataToExcel` します。</span><span class="sxs-lookup"><span data-stu-id="043a3-212">Delete the `writeDataToExcel` function.</span></span>

5. <span data-ttu-id="043a3-213">関数を削除 `writeDataToOutlook` します。</span><span class="sxs-lookup"><span data-stu-id="043a3-213">Delete the `writeDataToOutlook` function.</span></span>

6. <span data-ttu-id="043a3-214">関数を削除 `writeDataToPowerPoint` します。</span><span class="sxs-lookup"><span data-stu-id="043a3-214">Delete the `writeDataToPowerPoint` function.</span></span>

<span data-ttu-id="043a3-215">これらの変更を行った後で、この記事の「 [try a out](#try-it-out) 」セクションに進んで、更新されたアドインを試してみてください。</span><span class="sxs-lookup"><span data-stu-id="043a3-215">After you've made these changes, skip ahead to the [Try it out](#try-it-out) section of this article to try out your updated add-in.</span></span>

### <a name="changes-required-for-a-word-add-in-typescript"></a><span data-ttu-id="043a3-216">Word アドイン (TypeScript) に必要な変更</span><span class="sxs-lookup"><span data-stu-id="043a3-216">Changes required for a Word add-in (TypeScript)</span></span>

<span data-ttu-id="043a3-217">アドインが TypeScript を使用して作成された Word アドインである場合は、 **/src/taskpane/taskpane.ts**を開き、 `writeDataToOfficeDocument` 関数を見つけて、次の関数で置き換えます。</span><span class="sxs-lookup"><span data-stu-id="043a3-217">If your add-in is a Word add-in that was created with TypeScript, open **./src/taskpane/taskpane.ts**, find the `writeDataToOfficeDocument` function, and replace it with the following function:</span></span>

```typescript
export function writeDataToOfficeDocument(result: Object): Promise<any> {
  return Word.run(function(context) {
    let data: string[] = [];

    let itemNames: string[] = [];
    let oneDriveItems = result["value"];
    for (let item of oneDriveItems) {
      itemNames.push(item["name"]);
    };

    for (let i = 0; i < itemNames.length; i++) {
      if (itemNames[i] !== null) {
        data.push(itemNames[i]);
      }
    }

    const documentBody: Word.Body = context.document.body;
    for (let i = 0; i < data.length; i++) {
      if (data[i] !== null) {
        documentBody.insertParagraph(data[i], "End");
      }
    }
    return context.sync();
  });
}
```

<span data-ttu-id="043a3-218">これらの変更を行った後で、この記事の「[試行](#try-it-out)」セクションに進んで、更新されたアドインを試してみてください。</span><span class="sxs-lookup"><span data-stu-id="043a3-218">After you've made these changes, continue to the [Try it out](#try-it-out) section of this article to try out your updated add-in.</span></span>

## <a name="try-it-out"></a><span data-ttu-id="043a3-219">試してみる</span><span class="sxs-lookup"><span data-stu-id="043a3-219">Try it out</span></span>

<span data-ttu-id="043a3-220">アドインが Excel、Word、または PowerPoint アドインである場合は、次のセクションの手順を実行してみてください。アドインが Outlook アドインの場合は、代わりに[outlook](#outlook)セクションの手順を完了します。</span><span class="sxs-lookup"><span data-stu-id="043a3-220">If your add-in is an Excel, Word, or PowerPoint add-in, complete the steps in the following section to try it out. If your add-in is an Outlook add-in, complete the steps in the [Outlook](#outlook) section instead.</span></span>

### <a name="excel-word-and-powerpoint"></a><span data-ttu-id="043a3-221">Excel、Word、および PowerPoint</span><span class="sxs-lookup"><span data-stu-id="043a3-221">Excel, Word, and PowerPoint</span></span>

<span data-ttu-id="043a3-222">Excel、Word、または PowerPoint アドインを試すには、次の手順を実行します。</span><span class="sxs-lookup"><span data-stu-id="043a3-222">Complete the following steps to try out an Excel, Word, or PowerPoint add-in.</span></span>

1. <span data-ttu-id="043a3-223">プロジェクトのルートフォルダーで、次のコマンドを実行してプロジェクトをビルドし、ローカル web サーバーを起動して、以前に選択した Office クライアントアプリケーションでアドインをサイドロードします。</span><span class="sxs-lookup"><span data-stu-id="043a3-223">In the root folder of the project, run the following command to build the project, start the local web server, and sideload your add-in in the previously selected Office client application.</span></span>

    > [!NOTE]
    > <span data-ttu-id="043a3-224">開発の最中でも、Office アドインは HTTP ではなく HTTPS を使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="043a3-224">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="043a3-225">次のコマンドを実行した後に証明書をインストールするように求められた場合は、Yeoman ジェネレーターによって提供される証明書をインストールするプロンプトを受け入れます。</span><span class="sxs-lookup"><span data-stu-id="043a3-225">If you are prompted to install a certificate after you run the following command, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

    ```command&nbsp;line
    npm start
    ```

2. <span data-ttu-id="043a3-226">前のコマンド (つまり、Excel、Word、PowerPoint) を実行したときに開く Office クライアントアプリケーションで、アプリの[SSO の構成](sso-quickstart.md#configure-sso)時に Azure への接続に使用した microsoft 365 管理者アカウントと同じ microsoft 365 組織のメンバーであるユーザーを使用してサインインしていることを確認します。</span><span class="sxs-lookup"><span data-stu-id="043a3-226">In the Office client application that opens when you run the previous command (i.e., Excel, Word or PowerPoint), make sure that you're signed in with a user that's a member of the same Microsoft 365 organization as the Microsoft 365 administrator account that you used to connect to Azure while [configuring SSO](sso-quickstart.md#configure-sso) for the app.</span></span> <span data-ttu-id="043a3-227">これにより、SSO を正常に実行するための適切な条件が確立されます。</span><span class="sxs-lookup"><span data-stu-id="043a3-227">Doing so establishes the appropriate conditions for SSO to succeed.</span></span> 

3. <span data-ttu-id="043a3-228">Office クライアント アプリケーションで、[**ホーム**] タブを選択し、リボンの [**作業ウィンドウの表示**] ボタンをクリックして、アドインの作業ウィンドウを開きます。</span><span class="sxs-lookup"><span data-stu-id="043a3-228">In the Office client application, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span> <span data-ttu-id="043a3-229">次の画像は、Excel のこのボタンを示しています。</span><span class="sxs-lookup"><span data-stu-id="043a3-229">The following image shows this button in Excel.</span></span>

    ![Excel アドイン ボタン](../images/excel-quickstart-addin-3b.png)

4. <span data-ttu-id="043a3-231">作業ウィンドウの下部にある [ **OneDrive For business の読み取り**] ボタンをクリックして、SSO プロセスを開始します。</span><span class="sxs-lookup"><span data-stu-id="043a3-231">At the bottom of the task pane, choose the **Read my OneDrive for Business** button to initiate the SSO process.</span></span> 

5. <span data-ttu-id="043a3-232">アドインの代わりにアクセス許可を要求するダイアログ ウィンドウが表示される場合は、SSO はシナリオでサポートされず、代わりにアドインが別のユーザー認証方法に戻っていることを意味します。</span><span class="sxs-lookup"><span data-stu-id="043a3-232">If a dialog window appears to request permissions on behalf of the add-in, this means that SSO is not supported for your scenario and the add-in has instead fallen back to an alternate method of user authentication.</span></span> <span data-ttu-id="043a3-233">これは、テナント管理者が Microsoft Graph へのアクセスのためにアドインに同意を与えていない場合や、ユーザーが有効な Microsoft アカウントまたは Microsoft 365 の教育機関または勤務先のアカウントを使用して Office にサインインしていない場合に発生する可能性があります。</span><span class="sxs-lookup"><span data-stu-id="043a3-233">This may occur when the tenant administrator hasn't granted consent for the add-in to access Microsoft Graph, or when the user isn't signed into Office with a valid Microsoft Account or Microsoft 365 Education or Work account.</span></span> <span data-ttu-id="043a3-234">ダイアログ ウィンドウで [**同意する**] ボタンを選択して続行します。</span><span class="sxs-lookup"><span data-stu-id="043a3-234">Choose the **Accept** button in the dialog window to continue.</span></span>

    ![アクセス許可を要求するダイアログ](../images/sso-permissions-request.png)

    > [!NOTE]
    > <span data-ttu-id="043a3-236">ユーザーがこのアクセス許可の要求を受け入れると、今後再びプロンプトが表示されることはありません。</span><span class="sxs-lookup"><span data-stu-id="043a3-236">After a user accepts this permissions request, they won't be prompted again in the future.</span></span>

6. <span data-ttu-id="043a3-237">アドインは、サインインしているユーザーの OneDrive for Business からデータを読み取り、上位10個のファイルとフォルダーの名前をドキュメントに書き込みます。</span><span class="sxs-lookup"><span data-stu-id="043a3-237">The add-in reads data from the signed-in user's OneDrive for Business and writes the names of the top 10 files and folders to the document.</span></span> <span data-ttu-id="043a3-238">次の図は、Excel ワークシートに書き込まれるファイル名とフォルダー名の例を示しています。</span><span class="sxs-lookup"><span data-stu-id="043a3-238">The following image shows an example of file and folder names written to an Excel worksheet.</span></span>

    ![Excel ワークシートの OneDrive for Business 情報](../images/sso-onedrive-info-excel.png)

### <a name="outlook"></a><span data-ttu-id="043a3-240">Outlook</span><span class="sxs-lookup"><span data-stu-id="043a3-240">Outlook</span></span>

<span data-ttu-id="043a3-241">Outlook アドインを試すには、次の手順を実行します。</span><span class="sxs-lookup"><span data-stu-id="043a3-241">Complete the following steps to try out an Outlook add-in.</span></span>

1. <span data-ttu-id="043a3-242">プロジェクトのルートフォルダーで、次のコマンドを実行してプロジェクトをビルドし、ローカル web サーバーを開始します。</span><span class="sxs-lookup"><span data-stu-id="043a3-242">In the root folder of the project, run the following command to build the project and start the local web server.</span></span>

    > [!NOTE]
    > <span data-ttu-id="043a3-243">開発の最中でも、Office アドインは HTTP ではなく HTTPS を使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="043a3-243">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="043a3-244">次のコマンドを実行した後に証明書をインストールするように求められた場合は、Yeoman ジェネレーターによって提供される証明書をインストールするプロンプトを受け入れます。</span><span class="sxs-lookup"><span data-stu-id="043a3-244">If you are prompted to install a certificate after you run the following command, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

    ```command&nbsp;line
    npm start
    ```

2. <span data-ttu-id="043a3-245">「[テスト用に Outlook アドインをサイドロードする](/outlook/add-ins/sideload-outlook-add-ins-for-testing)」の手順に従って Outlook アドインをサイドロードします。</span><span class="sxs-lookup"><span data-stu-id="043a3-245">Follow the instructions in [Sideload Outlook add-ins for testing](/outlook/add-ins/sideload-outlook-add-ins-for-testing) to sideload the add-in in Outlook.</span></span> <span data-ttu-id="043a3-246">アプリの[SSO を構成](sso-quickstart.md#configure-sso)する際に Azure への接続に使用した microsoft 365 管理者アカウントと同じ microsoft 365 組織のメンバーであるユーザーを使用して、Outlook にサインインしていることを確認してください。</span><span class="sxs-lookup"><span data-stu-id="043a3-246">Make sure that you're signed in to Outlook with a user that's a member of the same Microsoft 365 organization as the Microsoft 365 administrator account that you used to connect to Azure while [configuring SSO](sso-quickstart.md#configure-sso) for the app.</span></span> <span data-ttu-id="043a3-247">これにより、SSO を正常に実行するための適切な条件が確立されます。</span><span class="sxs-lookup"><span data-stu-id="043a3-247">Doing so establishes the appropriate conditions for SSO to succeed.</span></span> 

3. <span data-ttu-id="043a3-248">Outlook で新しいメッセージを作成します。</span><span class="sxs-lookup"><span data-stu-id="043a3-248">In Outlook, compose a new message.</span></span>

4. <span data-ttu-id="043a3-249">[メッセージ作成] ウィンドウで、リボンの [**作業ウィンドウの表示**] ボタンを選択して、アドインの作業ウィンドウを開きます。</span><span class="sxs-lookup"><span data-stu-id="043a3-249">In the message compose window, choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Outlook アドイン ボタン](../images/outlook-sso-ribbon-button.png)

5. <span data-ttu-id="043a3-251">作業ウィンドウの下部にある [ **OneDrive For business の読み取り**] ボタンをクリックして、SSO プロセスを開始します。</span><span class="sxs-lookup"><span data-stu-id="043a3-251">At the bottom of the task pane, choose the **Read my OneDrive for Business** button to initiate the SSO process.</span></span> 

6. <span data-ttu-id="043a3-252">アドインの代わりにアクセス許可を要求するダイアログ ウィンドウが表示される場合は、SSO はシナリオでサポートされず、代わりにアドインが別のユーザー認証方法に戻っていることを意味します。</span><span class="sxs-lookup"><span data-stu-id="043a3-252">If a dialog window appears to request permissions on behalf of the add-in, this means that SSO is not supported for your scenario and the add-in has instead fallen back to an alternate method of user authentication.</span></span> <span data-ttu-id="043a3-253">これは、テナント管理者が Microsoft Graph へのアクセスのためにアドインに同意を与えていない場合や、ユーザーが有効な Microsoft アカウントまたは Microsoft 365 の教育機関または勤務先のアカウントを使用して Office にサインインしていない場合に発生する可能性があります。</span><span class="sxs-lookup"><span data-stu-id="043a3-253">This may occur when the tenant administrator hasn't granted consent for the add-in to access Microsoft Graph, or when the user isn't signed into Office with a valid Microsoft Account or Microsoft 365 Education or Work account.</span></span> <span data-ttu-id="043a3-254">ダイアログ ウィンドウで [**同意する**] ボタンを選択して続行します。</span><span class="sxs-lookup"><span data-stu-id="043a3-254">Choose the **Accept** button in the dialog window to continue.</span></span>

    ![アクセス許可を要求するダイアログ](../images/sso-permissions-request.png)

    > [!NOTE]
    > <span data-ttu-id="043a3-256">ユーザーがこのアクセス許可の要求を受け入れると、今後再びプロンプトが表示されることはありません。</span><span class="sxs-lookup"><span data-stu-id="043a3-256">After a user accepts this permissions request, they won't be prompted again in the future.</span></span>

7. <span data-ttu-id="043a3-257">アドインは、サインインしているユーザーの OneDrive for Business からデータを読み取り、上位10個のファイルとフォルダーの名前を電子メールメッセージの本文に書き込みます。</span><span class="sxs-lookup"><span data-stu-id="043a3-257">The add-in reads data from the signed-in user's OneDrive for Business and writes the names of the top 10 files and folders to the body of the email message.</span></span>

    ![Outlook メッセージの OneDrive for Business 情報](../images/sso-onedrive-info-outlook.png)

## <a name="next-steps"></a><span data-ttu-id="043a3-259">次の手順</span><span class="sxs-lookup"><span data-stu-id="043a3-259">Next steps</span></span>

<span data-ttu-id="043a3-260">これで、 [sso クイックスタート](sso-quickstart.md)で、[ごみ箱] ジェネレーターを使用して作成した sso を有効にしたアドインの機能をカスタマイズすることができました。</span><span class="sxs-lookup"><span data-stu-id="043a3-260">Congratulations, you've successfully customized the functionality of the SSO-enabled add-in that you created with the Yeoman generator in the [SSO quick start](sso-quickstart.md).</span></span> <span data-ttu-id="043a3-261">Yeoman ジェネレーターが自動的に完了した SSO の構成手順、および SSO プロセスを容易にするコードの詳細については、「[シングル サインオンを使用する Node.js Office アドインを作成する](../develop/create-sso-office-add-ins-nodejs.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="043a3-261">To learn more about SSO configuration steps that the Yeoman generator completed automatically, and the code that facilitates the SSO process, see the [Create a Node.js Office Add-in that uses single sign-on](../develop/create-sso-office-add-ins-nodejs.md) tutorial.</span></span>

## <a name="see-also"></a><span data-ttu-id="043a3-262">関連項目</span><span class="sxs-lookup"><span data-stu-id="043a3-262">See also</span></span>

- [<span data-ttu-id="043a3-263">Office アドインのシングル サインオンを有効化する</span><span class="sxs-lookup"><span data-stu-id="043a3-263">Enable single sign-on for Office Add-ins</span></span>](../develop/sso-in-office-add-ins.md)
- [<span data-ttu-id="043a3-264">シングル サインオン (SSO) のクイック スタート</span><span class="sxs-lookup"><span data-stu-id="043a3-264">Single sign-on (SSO) quick start</span></span>](sso-quickstart.md)
- [<span data-ttu-id="043a3-265">シングル サインオンを使用する Node.js Office アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="043a3-265">Create a Node.js Office Add-in that uses single sign-on</span></span>](../develop/create-sso-office-add-ins-nodejs.md)
- [<span data-ttu-id="043a3-266">シングル サインオン (SSO) のエラー メッセージのトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="043a3-266">Troubleshoot error messages for single sign-on (SSO)</span></span>](../develop/troubleshoot-sso-in-office-add-ins.md)
