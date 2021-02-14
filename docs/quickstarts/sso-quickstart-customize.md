---
title: Node.js SSO が有効なアドインをカスタマイズする
description: Yeoman ジェネレーターで作成した SSO 対応アドインのカスタマイズについて学習します。
ms.date: 02/01/2021
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 174df5e58e794b94b02025bd90a65f5ae8e26d44
ms.sourcegitcommit: fefc279b85e37463413b6b0e84c880d9ed5d7ac3
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/12/2021
ms.locfileid: "50234171"
---
# <a name="customize-your-nodejs-sso-enabled-add-in"></a><span data-ttu-id="f3c04-103">Node.js SSO が有効なアドインをカスタマイズする</span><span class="sxs-lookup"><span data-stu-id="f3c04-103">Customize your Node.js SSO-enabled add-in</span></span>

> [!IMPORTANT]
> <span data-ttu-id="f3c04-104">この記事は、シングル サインオン (SSO) クイック スタートを完了することによって作成される SSO 対応アドイン [に基になっています](sso-quickstart.md)。</span><span class="sxs-lookup"><span data-stu-id="f3c04-104">This article builds upon the SSO-enabled add-in that's created by completing the [single sign-on (SSO) quick start](sso-quickstart.md).</span></span> <span data-ttu-id="f3c04-105">この記事を読む前に、クイック スタートを完了してください。</span><span class="sxs-lookup"><span data-stu-id="f3c04-105">Please complete the quick start before reading this article.</span></span>

<span data-ttu-id="f3c04-106">[SSO クイック スタートは](sso-quickstart.md)、サインインしているユーザーのプロファイル情報を取得し、ドキュメントまたはメッセージに書き込む SSO 対応アドインを作成します。</span><span class="sxs-lookup"><span data-stu-id="f3c04-106">The [SSO quick start](sso-quickstart.md) creates an SSO-enabled add-in that gets the signed-in user's profile information and writes it to the document or message.</span></span> <span data-ttu-id="f3c04-107">この記事では、SSO クイック スタートで Yeoman ジェネレーターを使用して作成したアドインを更新し、さまざまなアクセス許可を必要とする新しい機能を追加するプロセスについて説明します。</span><span class="sxs-lookup"><span data-stu-id="f3c04-107">In this article, you'll walk through the process of updating the add-in that you created with the Yeoman generator in the SSO quick start, to add new functionality that requires different permissions.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="f3c04-108">前提条件</span><span class="sxs-lookup"><span data-stu-id="f3c04-108">Prerequisites</span></span>

- <span data-ttu-id="f3c04-109">SSO Office手順に従って作成したアドインの[1 つ。](sso-quickstart.md)</span><span class="sxs-lookup"><span data-stu-id="f3c04-109">An Office Add-in that you created by following the instructions in the [SSO quick start](sso-quickstart.md).</span></span>

- <span data-ttu-id="f3c04-110">Microsoft 365 サブスクリプションの OneDrive for Business に保存されているファイルとフォルダーが少なくともいくつかあります。</span><span class="sxs-lookup"><span data-stu-id="f3c04-110">At least a few files and folders stored on OneDrive for Business in your Microsoft 365 subscription.</span></span>

- <span data-ttu-id="f3c04-111">[Node.js](https://nodejs.org) (最新 [LTS](https://nodejs.org/about/releases) バージョン)。</span><span class="sxs-lookup"><span data-stu-id="f3c04-111">[Node.js](https://nodejs.org) (the latest [LTS](https://nodejs.org/about/releases) version).</span></span>

[!include[additional prerequisites](../includes/sso-tutorial-prereqs.md)]

## <a name="review-contents-of-the-project"></a><span data-ttu-id="f3c04-112">プロジェクトの内容を確認する</span><span class="sxs-lookup"><span data-stu-id="f3c04-112">Review contents of the project</span></span>

<span data-ttu-id="f3c04-113">最初に、Yeoman ジェネレーターで以前に作成したアドイン プロジェクトを [簡単に確認します](sso-quickstart.md)。</span><span class="sxs-lookup"><span data-stu-id="f3c04-113">Let's begin with a quick review of the add-in project that you previously [created with the Yeoman generator](sso-quickstart.md).</span></span>

> [!NOTE]
> <span data-ttu-id="f3c04-114">この記事で **.js** ファイル拡張子を使用してスクリプト ファイルを参照する場所では、プロジェクトが TypeScript で作成されている場合は、代わりに **.ts** ファイル拡張子を想定します。</span><span class="sxs-lookup"><span data-stu-id="f3c04-114">In places where this article references script files using **.js** file extension, assume the **.ts** file extension instead if your project was created with TypeScript.</span></span>

[!include[project structure for an SSO-enabled add-in created with the Yeoman generator](../includes/sso-yeoman-project-structure.md)]

## <a name="add-new-functionality"></a><span data-ttu-id="f3c04-115">新しい機能を追加する</span><span class="sxs-lookup"><span data-stu-id="f3c04-115">Add new functionality</span></span>

<span data-ttu-id="f3c04-116">SSO クイック スタートで作成したアドインは、Microsoft Graph を使用してサインインしているユーザーのプロファイル情報を取得し、その情報をドキュメントまたはメッセージに書き込みます。</span><span class="sxs-lookup"><span data-stu-id="f3c04-116">The add-in that you created with the SSO quick start uses Microsoft Graph to get the signed-in user's profile information and writes that information to the document or message.</span></span> <span data-ttu-id="f3c04-117">サインインしているユーザーの OneDrive for Business から上位 10 のファイルとフォルダーの名前を取得し、その情報をドキュメントまたはメッセージに書き込むアドインの機能を変更しましょう。</span><span class="sxs-lookup"><span data-stu-id="f3c04-117">Let's change the add-in's functionality such that it gets the names of the top 10 files and folders from the signed-in user's OneDrive for Business and writes that information to the document or message.</span></span> <span data-ttu-id="f3c04-118">この新機能を有効にする場合は、Azure でアプリのアクセス許可を更新し、アドイン プロジェクト内のコードを更新する必要があります。</span><span class="sxs-lookup"><span data-stu-id="f3c04-118">Enabling this new functionality requires updating app permissions in Azure and updating code within the add-in project.</span></span>

### <a name="update-app-permissions-in-azure"></a><span data-ttu-id="f3c04-119">Azure でアプリのアクセス許可を更新する</span><span class="sxs-lookup"><span data-stu-id="f3c04-119">Update app permissions in Azure</span></span>

<span data-ttu-id="f3c04-120">アドインがユーザーの OneDrive for Business のコンテンツを正常に読み取る前に、Azure のアプリ登録情報を適切なアクセス許可で更新する必要があります。</span><span class="sxs-lookup"><span data-stu-id="f3c04-120">Before the add-in can successfully read the contents of the user's OneDrive for Business, its app registration information in Azure must be updated with the appropriate permissions.</span></span> <span data-ttu-id="f3c04-121">次の手順を実行して、アプリに **Files.Read.All** アクセス許可を付与し **、User.Read** アクセス許可を取り消します。これは不要です。</span><span class="sxs-lookup"><span data-stu-id="f3c04-121">Complete the following steps to grant the app the **Files.Read.All** permission and revoke the **User.Read** permission, which is no longer needed.</span></span>

1. <span data-ttu-id="f3c04-122">Azure Portal [に移動し](https://ms.portal.azure.com/#home)**、Microsoft 365 管理者の資格情報を使用してサインインします**。</span><span class="sxs-lookup"><span data-stu-id="f3c04-122">Navigate to the [Azure portal](https://ms.portal.azure.com/#home) and **sign in using your Microsoft 365 administrator credentials**.</span></span>

2. <span data-ttu-id="f3c04-123">[アプリの登録 **] ページに移動** します。</span><span class="sxs-lookup"><span data-stu-id="f3c04-123">Navigate to the **App registrations** page.</span></span>
    > [!TIP]
    > <span data-ttu-id="f3c04-124">これを行うには、Azure のホームページで [アプリの登録] タイルを選択するか、ホーム ページの検索ボックスを使用してアプリの登録を検索して **選択** します。</span><span class="sxs-lookup"><span data-stu-id="f3c04-124">You can do this either by choosing the **App registrations** tile on the Azure home page or by using the search box on the home page to find and choose **App registrations**.</span></span>

3. <span data-ttu-id="f3c04-125">[ **アプリの登録] ページ** で、クイック スタート時に作成したアプリを選択します。</span><span class="sxs-lookup"><span data-stu-id="f3c04-125">On the **App registrations** page, choose the app that you created during the quick start.</span></span>
    > [!TIP]
    > <span data-ttu-id="f3c04-126">アプリ **の表示** 名は、Yeoman ジェネレーターを使用してプロジェクトを作成するときに指定したアドイン名と一致します。</span><span class="sxs-lookup"><span data-stu-id="f3c04-126">The **Display name** of the app will match the add-in name that you specified when you created the project with the Yeoman generator.</span></span>

4. <span data-ttu-id="f3c04-127">アプリの概要ページで、ページの左側にある [管理] 見出しの下にある **API** アクセス許可を選択します。</span><span class="sxs-lookup"><span data-stu-id="f3c04-127">From the app overview page, choose **API permissions** under the **Manage** heading on the left side of the page.</span></span>

5. <span data-ttu-id="f3c04-128">アクセス許可 **テーブルの User.Read** 行で、省略記号を選択し、表示されるメニューから [管理者の同意を取り消す] を選択します。</span><span class="sxs-lookup"><span data-stu-id="f3c04-128">In the **User.Read** row of the permissions table, choose the ellipsis and then select **Revoke admin consent** from the menu that appears.</span></span>

6. <span data-ttu-id="f3c04-129">表示される **プロンプトに応答** して [はい、削除] ボタンを選択します。</span><span class="sxs-lookup"><span data-stu-id="f3c04-129">Select the **Yes, remove** button in response to the prompt that's displayed.</span></span>

7. <span data-ttu-id="f3c04-130">アクセス許可 **テーブルの User.Read** 行で、省略記号を選択し、表示されるメニューから [アクセス許可の削除] を選択します。</span><span class="sxs-lookup"><span data-stu-id="f3c04-130">In the **User.Read** row of the permissions table, choose the ellipsis and then select **Remove permission** from the menu that appears.</span></span>

8. <span data-ttu-id="f3c04-131">表示される **プロンプトに応答** して [はい、削除] ボタンを選択します。</span><span class="sxs-lookup"><span data-stu-id="f3c04-131">Select the **Yes, remove** button in response to the prompt that's displayed.</span></span>

9. <span data-ttu-id="f3c04-132">**[アクセス許可の追加]** ボタンを選択します。</span><span class="sxs-lookup"><span data-stu-id="f3c04-132">Select the **Add a permission** button.</span></span>

10. <span data-ttu-id="f3c04-133">開いたパネルで **Microsoft Graph を選択し** 、[委任されたアクセス **許可] を選択します**。</span><span class="sxs-lookup"><span data-stu-id="f3c04-133">On the panel that opens choose **Microsoft Graph** and then choose **Delegated permissions**.</span></span>

11. <span data-ttu-id="f3c04-134">[API アクセス **許可の要求] パネルで、次の操作を** 行います。</span><span class="sxs-lookup"><span data-stu-id="f3c04-134">On the **Request API permissions** panel:</span></span>

    <span data-ttu-id="f3c04-135">a.</span><span class="sxs-lookup"><span data-stu-id="f3c04-135">a.</span></span> <span data-ttu-id="f3c04-136">[**ファイル] で\*\*\*\*、[Files.Read.All] を選択します**。</span><span class="sxs-lookup"><span data-stu-id="f3c04-136">Under **Files**, select **Files.Read.All**.</span></span>

    <span data-ttu-id="f3c04-137">b.</span><span class="sxs-lookup"><span data-stu-id="f3c04-137">b.</span></span> <span data-ttu-id="f3c04-138">パネルの **下部にある [アクセス許可の** 追加] ボタンを選択して、これらのアクセス許可の変更を保存します。</span><span class="sxs-lookup"><span data-stu-id="f3c04-138">Select the **Add permissions** button at the bottom of the panel to save these permissions changes.</span></span>

12. <span data-ttu-id="f3c04-139">[テナント **名] ボタンに管理者の同意を付与するを選択** します。</span><span class="sxs-lookup"><span data-stu-id="f3c04-139">Select the **Grant admin consent for [tenant name]** button.</span></span>

13. <span data-ttu-id="f3c04-140">表示される **プロンプトに** 応答して [はい] ボタンを選択します。</span><span class="sxs-lookup"><span data-stu-id="f3c04-140">Select the **Yes** button in response to the prompt that's displayed.</span></span>

### <a name="update-code-in-the-add-in-project"></a><span data-ttu-id="f3c04-141">アドイン プロジェクトのコードを更新する</span><span class="sxs-lookup"><span data-stu-id="f3c04-141">Update code in the add-in project</span></span>

<span data-ttu-id="f3c04-142">サインインしているユーザーの OneDrive for Business のコンテンツをアドインが読み取り可能にするには、次の必要があります。</span><span class="sxs-lookup"><span data-stu-id="f3c04-142">To enable the add-in to read contents of the signed-in user's OneDrive for Business, you'll need to:</span></span>

- <span data-ttu-id="f3c04-143">Microsoft Graph の URL、パラメーター、および必要なアクセス スコープを参照するコードを更新します。</span><span class="sxs-lookup"><span data-stu-id="f3c04-143">Update the code that references the Microsoft Graph URL, parameters, and required access scope.</span></span>

- <span data-ttu-id="f3c04-144">作業ウィンドウ UI を定義するコードを更新して、新しい機能を正確に説明します。</span><span class="sxs-lookup"><span data-stu-id="f3c04-144">Update the code that defines the task pane UI, so that it accurately describes the new functionality.</span></span>

- <span data-ttu-id="f3c04-145">Microsoft Graph からの応答を解析し、ドキュメントまたはメッセージに書き込むコードを更新します。</span><span class="sxs-lookup"><span data-stu-id="f3c04-145">Update the code that parses the response from Microsoft Graph and writes it to the document or message.</span></span>

<span data-ttu-id="f3c04-146">次の手順では、これらの更新プログラムについて説明します。</span><span class="sxs-lookup"><span data-stu-id="f3c04-146">The following steps describe these updates.</span></span>

### <a name="changes-required-for-any-type-of-add-in"></a><span data-ttu-id="f3c04-147">任意の種類のアドインに必要な変更</span><span class="sxs-lookup"><span data-stu-id="f3c04-147">Changes required for any type of add-in</span></span>

<span data-ttu-id="f3c04-148">アドインの次の手順を実行し、Microsoft Graph の URL、パラメーター、およびアクセス スコープを変更し、作業ウィンドウ UI を更新します。</span><span class="sxs-lookup"><span data-stu-id="f3c04-148">Complete the following steps for your add-in, to change the Microsoft Graph URL, parameters, and access scope, and update the task pane UI.</span></span> <span data-ttu-id="f3c04-149">これらの手順は、アドインがターゲットとするOfficeに関係なく同じです。</span><span class="sxs-lookup"><span data-stu-id="f3c04-149">These steps are the same, regardless of which Office application your add-in targets.</span></span>

1. <span data-ttu-id="f3c04-150">**./.ENV** ファイル:</span><span class="sxs-lookup"><span data-stu-id="f3c04-150">In the **./.ENV** file:</span></span>

    <span data-ttu-id="f3c04-151">a.</span><span class="sxs-lookup"><span data-stu-id="f3c04-151">a.</span></span> <span data-ttu-id="f3c04-152">次 `GRAPH_URL_SEGMENT=/me` のコードに置き換える: `GRAPH_URL_SEGMENT=/me/drive/root/children`</span><span class="sxs-lookup"><span data-stu-id="f3c04-152">Replace `GRAPH_URL_SEGMENT=/me` with the following: `GRAPH_URL_SEGMENT=/me/drive/root/children`</span></span>

    <span data-ttu-id="f3c04-153">b.</span><span class="sxs-lookup"><span data-stu-id="f3c04-153">b.</span></span> <span data-ttu-id="f3c04-154">次 `QUERY_PARAM_SEGMENT=` のコードに置き換える: `QUERY_PARAM_SEGMENT=?$select=name&$top=10`</span><span class="sxs-lookup"><span data-stu-id="f3c04-154">Replace `QUERY_PARAM_SEGMENT=` with the following: `QUERY_PARAM_SEGMENT=?$select=name&$top=10`</span></span>

    <span data-ttu-id="f3c04-155">c.</span><span class="sxs-lookup"><span data-stu-id="f3c04-155">c.</span></span> <span data-ttu-id="f3c04-156">次 `SCOPE=User.Read` のコードに置き換える: `SCOPE=Files.Read.All`</span><span class="sxs-lookup"><span data-stu-id="f3c04-156">Replace `SCOPE=User.Read` with the following: `SCOPE=Files.Read.All`</span></span>

2. <span data-ttu-id="f3c04-157">In **./manifest.xml,** find the line `<Scope>User.Read</Scope>` near the end of the file and replace it with the line `<Scope>Files.Read.All</Scope>` .</span><span class="sxs-lookup"><span data-stu-id="f3c04-157">In **./manifest.xml**, find the line `<Scope>User.Read</Scope>` near the end of the file and replace it with the line `<Scope>Files.Read.All</Scope>`.</span></span>

3. <span data-ttu-id="f3c04-158">**./src/helpers/fallbackauthdialog.js** (または TypeScript プロジェクトの **./src/helpers/fallbackauthdialog.ts)** で、次のように定義されている文字列を見つけて文字列に置き換えます `https://graph.microsoft.com/User.Read` `https://graph.microsoft.com/Files.Read.All` `requestObj` 。</span><span class="sxs-lookup"><span data-stu-id="f3c04-158">In **./src/helpers/fallbackauthdialog.js** (or in **./src/helpers/fallbackauthdialog.ts** for a TypeScript project), find the string `https://graph.microsoft.com/User.Read` and replace it with the string `https://graph.microsoft.com/Files.Read.All`, such that `requestObj` is defined as follows:</span></span>

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

4. <span data-ttu-id="f3c04-159">**./src/taskpane/taskpane.html** で要素を見つけ、その要素内のテキストを更新して、アドインの新機能を `<section class="ms-firstrun-instructionstep__header">` 説明します。</span><span class="sxs-lookup"><span data-stu-id="f3c04-159">In **./src/taskpane/taskpane.html**, find the element `<section class="ms-firstrun-instructionstep__header">` and update the text within that element to describe the add-in's new functionality.</span></span>

    ```html
    <section class="ms-firstrun-instructionstep__header">
        <h2 class="ms-font-m">This add-in demonstrates how to use single sign-on by making a call to Microsoft
            Graph to read content from OneDrive for Business.</h2>
        <div class="ms-firstrun-instructionstep__header--image"></div>
    </section>
    ```

5. <span data-ttu-id="f3c04-160">in **./src/taskpane/taskpane.html,** find and replace both occurrences of the `Get My User Profile Information` string with the string `Read my OneDrive for Business` .</span><span class="sxs-lookup"><span data-stu-id="f3c04-160">In **./src/taskpane/taskpane.html**, find and replace both occurrences of the string `Get My User Profile Information` with the string `Read my OneDrive for Business`.</span></span>

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

6. <span data-ttu-id="f3c04-161">in **./src/taskpane/taskpane.html,** find and replace the string `Your user profile information will be displayed in the document.` with the string `The names of the top 10 files and folders in your OneDrive for Business will be displayed in the document or message.` .</span><span class="sxs-lookup"><span data-stu-id="f3c04-161">In **./src/taskpane/taskpane.html**, find and replace the string `Your user profile information will be displayed in the document.` with the string `The names of the top 10 files and folders in your OneDrive for Business will be displayed in the document or message.`.</span></span>

    ```html
    <li class="ms-ListItem">
        <span class="ms-ListItem-primaryText">The names of the top 10 files and folders in your OneDrive for Business will be displayed in the document or message.</span>
        <div class="clearfix"></div>
    </li>
    ```

7. <span data-ttu-id="f3c04-162">Microsoft Graph からの応答を解析し、ドキュメントまたはメッセージに書き込むコードを、アドインの種類に対応するセクションのガイダンスに従って更新します。</span><span class="sxs-lookup"><span data-stu-id="f3c04-162">Update the code that parses the response from Microsoft Graph and writes it to the document or message, by following guidance in the section that corresponds to your type of add-in:</span></span>

    - [<span data-ttu-id="f3c04-163">Excel アドインに必要な変更 (JavaScript)</span><span class="sxs-lookup"><span data-stu-id="f3c04-163">Changes required for an Excel add-in (JavaScript)</span></span>](#changes-required-for-an-excel-add-in-javascript)
    - [<span data-ttu-id="f3c04-164">Excel アドインに必要な変更 (TypeScript)</span><span class="sxs-lookup"><span data-stu-id="f3c04-164">Changes required for an Excel add-in (TypeScript)</span></span>](#changes-required-for-an-excel-add-in-typescript)
    - [<span data-ttu-id="f3c04-165">Outlook アドインに必要な変更 (JavaScript)</span><span class="sxs-lookup"><span data-stu-id="f3c04-165">Changes required for an Outlook add-in (JavaScript)</span></span>](#changes-required-for-an-outlook-add-in-javascript)
    - [<span data-ttu-id="f3c04-166">Outlook アドインに必要な変更 (TypeScript)</span><span class="sxs-lookup"><span data-stu-id="f3c04-166">Changes required for an Outlook add-in (TypeScript)</span></span>](#changes-required-for-an-outlook-add-in-typescript)
    - [<span data-ttu-id="f3c04-167">PowerPoint アドインに必要な変更 (JavaScript)</span><span class="sxs-lookup"><span data-stu-id="f3c04-167">Changes required for a PowerPoint add-in (JavaScript)</span></span>](#changes-required-for-a-powerpoint-add-in-javascript)
    - [<span data-ttu-id="f3c04-168">PowerPoint アドインに必要な変更 (TypeScript)</span><span class="sxs-lookup"><span data-stu-id="f3c04-168">Changes required for a PowerPoint add-in (TypeScript)</span></span>](#changes-required-for-a-powerpoint-add-in-typescript)
    - [<span data-ttu-id="f3c04-169">Word アドインに必要な変更 (JavaScript)</span><span class="sxs-lookup"><span data-stu-id="f3c04-169">Changes required for a Word add-in (JavaScript)</span></span>](#changes-required-for-a-word-add-in-javascript)
    - [<span data-ttu-id="f3c04-170">Word アドインに必要な変更 (TypeScript)</span><span class="sxs-lookup"><span data-stu-id="f3c04-170">Changes required for a Word add-in (TypeScript)</span></span>](#changes-required-for-a-word-add-in-typescript)

### <a name="changes-required-for-an-excel-add-in-javascript"></a><span data-ttu-id="f3c04-171">Excel アドインに必要な変更 (JavaScript)</span><span class="sxs-lookup"><span data-stu-id="f3c04-171">Changes required for an Excel add-in (JavaScript)</span></span>

<span data-ttu-id="f3c04-172">アドインが JavaScript で作成された Excel アドインである場合は **、./src/helpers/documentHelper.js** で次の変更を行います。</span><span class="sxs-lookup"><span data-stu-id="f3c04-172">If your add-in is an Excel add-in that was created with JavaScript, make the following changes in **./src/helpers/documentHelper.js**:</span></span>

1. <span data-ttu-id="f3c04-173">関数を `writeDataToOfficeDocument` 検索し、次の関数に置き換える。</span><span class="sxs-lookup"><span data-stu-id="f3c04-173">Find the `writeDataToOfficeDocument` function and replace it with the following function:</span></span>

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

2. <span data-ttu-id="f3c04-174">関数を `filterUserProfileInfo` 検索し、次の関数に置き換える。</span><span class="sxs-lookup"><span data-stu-id="f3c04-174">Find the `filterUserProfileInfo` function and replace it with the following function:</span></span>

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

3. <span data-ttu-id="f3c04-175">関数を `writeDataToExcel` 検索し、次の関数に置き換える。</span><span class="sxs-lookup"><span data-stu-id="f3c04-175">Find the `writeDataToExcel` function and replace it with the following function:</span></span>

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

4. <span data-ttu-id="f3c04-176">関数を削除 `writeDataToOutlook` します。</span><span class="sxs-lookup"><span data-stu-id="f3c04-176">Delete the `writeDataToOutlook` function.</span></span>

5. <span data-ttu-id="f3c04-177">関数を削除 `writeDataToPowerPoint` します。</span><span class="sxs-lookup"><span data-stu-id="f3c04-177">Delete the `writeDataToPowerPoint` function.</span></span>

6. <span data-ttu-id="f3c04-178">関数を削除 `writeDataToWord` します。</span><span class="sxs-lookup"><span data-stu-id="f3c04-178">Delete the `writeDataToWord` function.</span></span>

<span data-ttu-id="f3c04-179">これらの変更を行った後、この記事の「試[](#try-it-out)してみる」セクションに進み、更新されたアドインを試してください。</span><span class="sxs-lookup"><span data-stu-id="f3c04-179">After you've made these changes, skip ahead to the [Try it out](#try-it-out) section of this article to try out your updated add-in.</span></span>

### <a name="changes-required-for-an-excel-add-in-typescript"></a><span data-ttu-id="f3c04-180">Excel アドインに必要な変更 (TypeScript)</span><span class="sxs-lookup"><span data-stu-id="f3c04-180">Changes required for an Excel add-in (TypeScript)</span></span>

<span data-ttu-id="f3c04-181">アドインが TypeScript で作成された Excel アドインである場合は **、./src/taskpane/taskpane.ts** を開き、関数を見つけて、次の関数に置き換 `writeDataToOfficeDocument` える必要があります。</span><span class="sxs-lookup"><span data-stu-id="f3c04-181">If your add-in is an Excel add-in that was created with TypeScript, open **./src/taskpane/taskpane.ts**, find the `writeDataToOfficeDocument` function, and replace it with the following function:</span></span>

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

<span data-ttu-id="f3c04-182">これらの変更を行った後、この記事の「試[](#try-it-out)してみる」セクションに進み、更新されたアドインを試してください。</span><span class="sxs-lookup"><span data-stu-id="f3c04-182">After you've made these changes, skip ahead to the [Try it out](#try-it-out) section of this article to try out your updated add-in.</span></span>

### <a name="changes-required-for-an-outlook-add-in-javascript"></a><span data-ttu-id="f3c04-183">Outlook アドインに必要な変更 (JavaScript)</span><span class="sxs-lookup"><span data-stu-id="f3c04-183">Changes required for an Outlook add-in (JavaScript)</span></span>

<span data-ttu-id="f3c04-184">アドインが JavaScript で作成された Outlook アドインである場合は **、./src/helpers/documentHelper.js** で次の変更を行います。</span><span class="sxs-lookup"><span data-stu-id="f3c04-184">If your add-in is an Outlook add-in that was created with JavaScript, make the following changes in **./src/helpers/documentHelper.js**:</span></span>

1. <span data-ttu-id="f3c04-185">関数を `writeDataToOfficeDocument` 検索し、次の関数に置き換える。</span><span class="sxs-lookup"><span data-stu-id="f3c04-185">Find the `writeDataToOfficeDocument` function and replace it with the following function:</span></span>

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

2. <span data-ttu-id="f3c04-186">関数を `filterUserProfileInfo` 検索し、次の関数に置き換える。</span><span class="sxs-lookup"><span data-stu-id="f3c04-186">Find the `filterUserProfileInfo` function and replace it with the following function:</span></span>

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

3. <span data-ttu-id="f3c04-187">関数を `writeDataToOutlook` 検索し、次の関数に置き換える。</span><span class="sxs-lookup"><span data-stu-id="f3c04-187">Find the `writeDataToOutlook` function and replace it with the following function:</span></span>

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

4. <span data-ttu-id="f3c04-188">関数を削除 `writeDataToExcel` します。</span><span class="sxs-lookup"><span data-stu-id="f3c04-188">Delete the `writeDataToExcel` function.</span></span>

5. <span data-ttu-id="f3c04-189">関数を削除 `writeDataToPowerPoint` します。</span><span class="sxs-lookup"><span data-stu-id="f3c04-189">Delete the `writeDataToPowerPoint` function.</span></span>

6. <span data-ttu-id="f3c04-190">関数を削除 `writeDataToWord` します。</span><span class="sxs-lookup"><span data-stu-id="f3c04-190">Delete the `writeDataToWord` function.</span></span>

<span data-ttu-id="f3c04-191">これらの変更を行った後、この記事の「試[](#try-it-out)してみる」セクションに進み、更新されたアドインを試してください。</span><span class="sxs-lookup"><span data-stu-id="f3c04-191">After you've made these changes, skip ahead to the [Try it out](#try-it-out) section of this article to try out your updated add-in.</span></span>

### <a name="changes-required-for-an-outlook-add-in-typescript"></a><span data-ttu-id="f3c04-192">Outlook アドインに必要な変更 (TypeScript)</span><span class="sxs-lookup"><span data-stu-id="f3c04-192">Changes required for an Outlook add-in (TypeScript)</span></span>

<span data-ttu-id="f3c04-193">アドインが TypeScript で作成された Outlook アドインである場合は **、./src/taskpane/taskpane.ts** を開き、関数を見つけて、次の関数に置き換 `writeDataToOfficeDocument` える必要があります。</span><span class="sxs-lookup"><span data-stu-id="f3c04-193">If your add-in is an Outlook add-in that was created with TypeScript, open **./src/taskpane/taskpane.ts**, find the `writeDataToOfficeDocument` function, and replace it with the following function:</span></span>

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

<span data-ttu-id="f3c04-194">これらの変更を行った後、この記事の「試[](#try-it-out)してみる」セクションに進み、更新されたアドインを試してください。</span><span class="sxs-lookup"><span data-stu-id="f3c04-194">After you've made these changes, skip ahead to the [Try it out](#try-it-out) section of this article to try out your updated add-in.</span></span>

### <a name="changes-required-for-a-powerpoint-add-in-javascript"></a><span data-ttu-id="f3c04-195">PowerPoint アドインに必要な変更 (JavaScript)</span><span class="sxs-lookup"><span data-stu-id="f3c04-195">Changes required for a PowerPoint add-in (JavaScript)</span></span>

<span data-ttu-id="f3c04-196">アドインが JavaScript で作成された PowerPoint アドインである場合は **、./src/helpers/documentHelper.js** で次の変更を行います。</span><span class="sxs-lookup"><span data-stu-id="f3c04-196">If your add-in is a PowerPoint add-in that was created with JavaScript, make the following changes in **./src/helpers/documentHelper.js**:</span></span>

1. <span data-ttu-id="f3c04-197">関数を `writeDataToOfficeDocument` 検索し、次の関数に置き換える。</span><span class="sxs-lookup"><span data-stu-id="f3c04-197">Find the `writeDataToOfficeDocument` function and replace it with the following function:</span></span>

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

2. <span data-ttu-id="f3c04-198">関数を `filterUserProfileInfo` 検索し、次の関数に置き換える。</span><span class="sxs-lookup"><span data-stu-id="f3c04-198">Find the `filterUserProfileInfo` function and replace it with the following function:</span></span>

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

3. <span data-ttu-id="f3c04-199">関数を `writeDataToPowerPoint` 検索し、次の関数に置き換える。</span><span class="sxs-lookup"><span data-stu-id="f3c04-199">Find the `writeDataToPowerPoint` function and replace it with the following function:</span></span>

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

4. <span data-ttu-id="f3c04-200">関数を削除 `writeDataToExcel` します。</span><span class="sxs-lookup"><span data-stu-id="f3c04-200">Delete the `writeDataToExcel` function.</span></span>

5. <span data-ttu-id="f3c04-201">関数を削除 `writeDataToOutlook` します。</span><span class="sxs-lookup"><span data-stu-id="f3c04-201">Delete the `writeDataToOutlook` function.</span></span>

6. <span data-ttu-id="f3c04-202">関数を削除 `writeDataToWord` します。</span><span class="sxs-lookup"><span data-stu-id="f3c04-202">Delete the `writeDataToWord` function.</span></span>

<span data-ttu-id="f3c04-203">これらの変更を行った後、この記事の「試[](#try-it-out)してみる」セクションに進み、更新されたアドインを試してください。</span><span class="sxs-lookup"><span data-stu-id="f3c04-203">After you've made these changes, skip ahead to the [Try it out](#try-it-out) section of this article to try out your updated add-in.</span></span>

### <a name="changes-required-for-a-powerpoint-add-in-typescript"></a><span data-ttu-id="f3c04-204">PowerPoint アドインに必要な変更 (TypeScript)</span><span class="sxs-lookup"><span data-stu-id="f3c04-204">Changes required for a PowerPoint add-in (TypeScript)</span></span>

<span data-ttu-id="f3c04-205">アドインが TypeScript で作成された PowerPoint アドインである場合は **、./src/taskpane/taskpane.ts** を開き、関数を見つけて、次の関数に置き換 `writeDataToOfficeDocument` える必要があります。</span><span class="sxs-lookup"><span data-stu-id="f3c04-205">If your add-in is a PowerPoint add-in that was created with TypeScript, open **./src/taskpane/taskpane.ts**, find the `writeDataToOfficeDocument` function, and replace it with the following function:</span></span>

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

<span data-ttu-id="f3c04-206">これらの変更を行った後、この記事の「試[](#try-it-out)してみる」セクションに進み、更新されたアドインを試してください。</span><span class="sxs-lookup"><span data-stu-id="f3c04-206">After you've made these changes, skip ahead to the [Try it out](#try-it-out) section of this article to try out your updated add-in.</span></span>

### <a name="changes-required-for-a-word-add-in-javascript"></a><span data-ttu-id="f3c04-207">Word アドインに必要な変更 (JavaScript)</span><span class="sxs-lookup"><span data-stu-id="f3c04-207">Changes required for a Word add-in (JavaScript)</span></span>

<span data-ttu-id="f3c04-208">アドインが JavaScript で作成された Word アドインである場合は **、./src/helpers/documentHelper.js** で次の変更を行います。</span><span class="sxs-lookup"><span data-stu-id="f3c04-208">If your add-in is a Word add-in that was created with JavaScript, make the following changes in **./src/helpers/documentHelper.js**:</span></span>

1. <span data-ttu-id="f3c04-209">関数を `writeDataToOfficeDocument` 検索し、次の関数に置き換える。</span><span class="sxs-lookup"><span data-stu-id="f3c04-209">Find the `writeDataToOfficeDocument` function and replace it with the following function:</span></span>

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

2. <span data-ttu-id="f3c04-210">関数を `filterUserProfileInfo` 検索し、次の関数に置き換える。</span><span class="sxs-lookup"><span data-stu-id="f3c04-210">Find the `filterUserProfileInfo` function and replace it with the following function:</span></span>

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

3. <span data-ttu-id="f3c04-211">関数を `writeDataToWord` 検索し、次の関数に置き換える。</span><span class="sxs-lookup"><span data-stu-id="f3c04-211">Find the `writeDataToWord` function and replace it with the following function:</span></span>

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

4. <span data-ttu-id="f3c04-212">関数を削除 `writeDataToExcel` します。</span><span class="sxs-lookup"><span data-stu-id="f3c04-212">Delete the `writeDataToExcel` function.</span></span>

5. <span data-ttu-id="f3c04-213">関数を削除 `writeDataToOutlook` します。</span><span class="sxs-lookup"><span data-stu-id="f3c04-213">Delete the `writeDataToOutlook` function.</span></span>

6. <span data-ttu-id="f3c04-214">関数を削除 `writeDataToPowerPoint` します。</span><span class="sxs-lookup"><span data-stu-id="f3c04-214">Delete the `writeDataToPowerPoint` function.</span></span>

<span data-ttu-id="f3c04-215">これらの変更を行った後、この記事の「試[](#try-it-out)してみる」セクションに進み、更新されたアドインを試してください。</span><span class="sxs-lookup"><span data-stu-id="f3c04-215">After you've made these changes, skip ahead to the [Try it out](#try-it-out) section of this article to try out your updated add-in.</span></span>

### <a name="changes-required-for-a-word-add-in-typescript"></a><span data-ttu-id="f3c04-216">Word アドインに必要な変更 (TypeScript)</span><span class="sxs-lookup"><span data-stu-id="f3c04-216">Changes required for a Word add-in (TypeScript)</span></span>

<span data-ttu-id="f3c04-217">アドインが TypeScript で作成された Word アドインである場合は **、./src/taskpane/taskpane.ts** を開き、関数を見つけて、次の関数に置き換 `writeDataToOfficeDocument` える必要があります。</span><span class="sxs-lookup"><span data-stu-id="f3c04-217">If your add-in is a Word add-in that was created with TypeScript, open **./src/taskpane/taskpane.ts**, find the `writeDataToOfficeDocument` function, and replace it with the following function:</span></span>

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

<span data-ttu-id="f3c04-218">これらの変更を行った後、この記事の「[](#try-it-out)試してみる」セクションに進み、更新されたアドインを試してください。</span><span class="sxs-lookup"><span data-stu-id="f3c04-218">After you've made these changes, continue to the [Try it out](#try-it-out) section of this article to try out your updated add-in.</span></span>

## <a name="try-it-out"></a><span data-ttu-id="f3c04-219">試してみる</span><span class="sxs-lookup"><span data-stu-id="f3c04-219">Try it out</span></span>

<span data-ttu-id="f3c04-220">アドインが Excel、Word、または PowerPoint のアドインである場合は、次のセクションの手順を実行して試してください。アドインが Outlook アドインの場合は、代わりに [Outlook](#outlook) セクションの手順を完了します。</span><span class="sxs-lookup"><span data-stu-id="f3c04-220">If your add-in is an Excel, Word, or PowerPoint add-in, complete the steps in the following section to try it out. If your add-in is an Outlook add-in, complete the steps in the [Outlook](#outlook) section instead.</span></span>

### <a name="excel-word-and-powerpoint"></a><span data-ttu-id="f3c04-221">Excel、Word、および PowerPoint</span><span class="sxs-lookup"><span data-stu-id="f3c04-221">Excel, Word, and PowerPoint</span></span>

<span data-ttu-id="f3c04-222">Excel、Word、または PowerPoint アドインを試すには、次の手順を実行します。</span><span class="sxs-lookup"><span data-stu-id="f3c04-222">Complete the following steps to try out an Excel, Word, or PowerPoint add-in.</span></span>

1. <span data-ttu-id="f3c04-223">プロジェクトのルート フォルダーで、次のコマンドを実行してプロジェクトをビルドし、ローカル Web サーバーを起動し、以前に選択した Office クライアント アプリケーションでアドインをサイドロードします。</span><span class="sxs-lookup"><span data-stu-id="f3c04-223">In the root folder of the project, run the following command to build the project, start the local web server, and sideload your add-in in the previously selected Office client application.</span></span>

    > [!NOTE]
    > <span data-ttu-id="f3c04-224">開発の最中でも、OfficeアドインはHTTPではなくHTTPSを使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="f3c04-224">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="f3c04-225">次のコマンドを実行した後に証明書をインストールするように求められた場合は、Yeoman ジェネレーターによって提供される証明書をインストールするプロンプトを受け入れます。</span><span class="sxs-lookup"><span data-stu-id="f3c04-225">If you are prompted to install a certificate after you run the following command, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

    ```command&nbsp;line
    npm start
    ```

2. <span data-ttu-id="f3c04-226">前のコマンド (Excel、Word、PowerPoint など) を実行するときに開く Office クライアント アプリケーションで、アプリの [SSO](sso-quickstart.md#configure-sso) の構成中に Azure への接続に使用した Microsoft 365 管理者アカウントと同じ Microsoft 365 組織のメンバーであるユーザーとサインインしている必要があります。</span><span class="sxs-lookup"><span data-stu-id="f3c04-226">In the Office client application that opens when you run the previous command (i.e., Excel, Word or PowerPoint), make sure that you're signed in with a user that's a member of the same Microsoft 365 organization as the Microsoft 365 administrator account that you used to connect to Azure while [configuring SSO](sso-quickstart.md#configure-sso) for the app.</span></span> <span data-ttu-id="f3c04-227">これにより、SSO を正常に実行するための適切な条件が確立されます。</span><span class="sxs-lookup"><span data-stu-id="f3c04-227">Doing so establishes the appropriate conditions for SSO to succeed.</span></span> 

3. <span data-ttu-id="f3c04-228">Office クライアント アプリケーションで、[**ホーム**] タブを選択し、リボンの [**作業ウィンドウの表示**] ボタンをクリックして、アドインの作業ウィンドウを開きます。</span><span class="sxs-lookup"><span data-stu-id="f3c04-228">In the Office client application, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span> <span data-ttu-id="f3c04-229">次の画像は、Excel のこのボタンを示しています。</span><span class="sxs-lookup"><span data-stu-id="f3c04-229">The following image shows this button in Excel.</span></span>

    ![Excel リボンで強調表示されたアドイン ボタンを示すスクリーンショット](../images/excel-quickstart-addin-3b.png)

4. <span data-ttu-id="f3c04-231">作業ウィンドウの下部にある **[OneDrive for Business** の読み取り] ボタンを選択して、SSO プロセスを開始します。</span><span class="sxs-lookup"><span data-stu-id="f3c04-231">At the bottom of the task pane, choose the **Read my OneDrive for Business** button to initiate the SSO process.</span></span>

5. <span data-ttu-id="f3c04-232">アドインの代わりにアクセス許可を要求するダイアログ ウィンドウが表示される場合は、SSO はシナリオでサポートされず、代わりにアドインが別のユーザー認証方法に戻っていることを意味します。</span><span class="sxs-lookup"><span data-stu-id="f3c04-232">If a dialog window appears to request permissions on behalf of the add-in, this means that SSO is not supported for your scenario and the add-in has instead fallen back to an alternate method of user authentication.</span></span> <span data-ttu-id="f3c04-233">これは、アドインが Microsoft Graph にアクセスすることに対してテナント管理者が同意を与えていない場合、または、ユーザーが有効な Microsoft アカウント、Microsoft 365 Education または職場アカウントで Office にサインインしていない場合に発生することがあります。</span><span class="sxs-lookup"><span data-stu-id="f3c04-233">This may occur when the tenant administrator hasn't granted consent for the add-in to access Microsoft Graph, or when the user isn't signed into Office with a valid Microsoft account or Microsoft 365 Education or Work account.</span></span> <span data-ttu-id="f3c04-234">ダイアログ ウィンドウで [**同意する**] ボタンを選択して続行します。</span><span class="sxs-lookup"><span data-stu-id="f3c04-234">Choose the **Accept** button in the dialog window to continue.</span></span>

    ![[承認] ボタンが強調表示された [アクセス許可] 要求ダイアログを示すスクリーンショット](../images/sso-permissions-request.png)

    > [!NOTE]
    > <span data-ttu-id="f3c04-236">ユーザーがこのアクセス許可の要求を受け入れると、今後再びプロンプトが表示されることはありません。</span><span class="sxs-lookup"><span data-stu-id="f3c04-236">After a user accepts this permissions request, they won't be prompted again in the future.</span></span>

6. <span data-ttu-id="f3c04-237">アドインは、サインインしているユーザーの OneDrive for Business からデータを読み取り、上位 10 のファイルとフォルダーの名前をドキュメントに書き込みます。</span><span class="sxs-lookup"><span data-stu-id="f3c04-237">The add-in reads data from the signed-in user's OneDrive for Business and writes the names of the top 10 files and folders to the document.</span></span> <span data-ttu-id="f3c04-238">次の図は、Excel ワークシートに書き込まれたファイル名とフォルダー名の例を示しています。</span><span class="sxs-lookup"><span data-stu-id="f3c04-238">The following image shows an example of file and folder names written to an Excel worksheet.</span></span>

    ![Excel ワークシートの OneDrive for Business 情報を示すスクリーンショット](../images/sso-onedrive-info-excel.png)

### <a name="outlook"></a><span data-ttu-id="f3c04-240">Outlook</span><span class="sxs-lookup"><span data-stu-id="f3c04-240">Outlook</span></span>

<span data-ttu-id="f3c04-241">Outlook アドインを試すには、次の手順を実行します。</span><span class="sxs-lookup"><span data-stu-id="f3c04-241">Complete the following steps to try out an Outlook add-in.</span></span>

1. <span data-ttu-id="f3c04-242">プロジェクトのルート フォルダーで、次のコマンドを実行してプロジェクトをビルドし、ローカル Web サーバーを起動し、アドインをサイドロードします。</span><span class="sxs-lookup"><span data-stu-id="f3c04-242">In the root folder of the project, run the following command to build the project, start the local web server, and sideload your add-in.</span></span> 

    > [!NOTE]
    > <span data-ttu-id="f3c04-243">開発の最中でも、OfficeアドインはHTTPではなくHTTPSを使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="f3c04-243">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="f3c04-244">次のコマンドを実行した後に証明書をインストールするように求められた場合は、Yeoman ジェネレーターによって提供される証明書をインストールするプロンプトを受け入れます。</span><span class="sxs-lookup"><span data-stu-id="f3c04-244">If you are prompted to install a certificate after you run the following command, accept the prompt to install the certificate that the Yeoman generator provides.</span></span> <span data-ttu-id="f3c04-245">変更を行うには、管理者としてコマンド プロンプトまたはターミナルを実行する必要がある場合もあります。</span><span class="sxs-lookup"><span data-stu-id="f3c04-245">You may also have to run your command prompt or terminal as an administrator for the changes to be made.</span></span>

    ```command&nbsp;line
    npm start
    ```

2. <span data-ttu-id="f3c04-246">アプリの [SSO](sso-quickstart.md#configure-sso) の構成中に Azure への接続に使用した Microsoft 365 管理者アカウントと同じ Microsoft 365 組織のメンバーであるユーザーと Outlook にサインインしている必要があります。</span><span class="sxs-lookup"><span data-stu-id="f3c04-246">Make sure that you're signed in to Outlook with a user that's a member of the same Microsoft 365 organization as the Microsoft 365 administrator account that you used to connect to Azure while [configuring SSO](sso-quickstart.md#configure-sso) for the app.</span></span> <span data-ttu-id="f3c04-247">これにより、SSO を正常に実行するための適切な条件が確立されます。</span><span class="sxs-lookup"><span data-stu-id="f3c04-247">Doing so establishes the appropriate conditions for SSO to succeed.</span></span>

3. <span data-ttu-id="f3c04-248">Outlook で新しいメッセージを作成します。</span><span class="sxs-lookup"><span data-stu-id="f3c04-248">In Outlook, compose a new message.</span></span>

4. <span data-ttu-id="f3c04-249">[メッセージ作成] ウィンドウで、リボンの [**作業ウィンドウの表示**] ボタンを選択して、アドインの作業ウィンドウを開きます。</span><span class="sxs-lookup"><span data-stu-id="f3c04-249">In the message compose window, choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Outlook の [メッセージの作成] ウィンドウの [強調表示されたアドイン] リボンボタンを示すスクリーン ショット](../images/outlook-sso-ribbon-button.png)

5. <span data-ttu-id="f3c04-251">作業ウィンドウの下部にある **[OneDrive for Business** の読み取り] ボタンを選択して、SSO プロセスを開始します。</span><span class="sxs-lookup"><span data-stu-id="f3c04-251">At the bottom of the task pane, choose the **Read my OneDrive for Business** button to initiate the SSO process.</span></span>

6. <span data-ttu-id="f3c04-252">アドインの代わりにアクセス許可を要求するダイアログ ウィンドウが表示される場合は、SSO はシナリオでサポートされず、代わりにアドインが別のユーザー認証方法に戻っていることを意味します。</span><span class="sxs-lookup"><span data-stu-id="f3c04-252">If a dialog window appears to request permissions on behalf of the add-in, this means that SSO is not supported for your scenario and the add-in has instead fallen back to an alternate method of user authentication.</span></span> <span data-ttu-id="f3c04-253">これは、アドインが Microsoft Graph にアクセスすることに対してテナント管理者が同意を与えていない場合、または、ユーザーが有効な Microsoft アカウント、Microsoft 365 Education または職場アカウントで Office にサインインしていない場合に発生することがあります。</span><span class="sxs-lookup"><span data-stu-id="f3c04-253">This may occur when the tenant administrator hasn't granted consent for the add-in to access Microsoft Graph, or when the user isn't signed into Office with a valid Microsoft account or Microsoft 365 Education or Work account.</span></span> <span data-ttu-id="f3c04-254">ダイアログ ウィンドウで [**同意する**] ボタンを選択して続行します。</span><span class="sxs-lookup"><span data-stu-id="f3c04-254">Choose the **Accept** button in the dialog window to continue.</span></span>

    ![[承認] ボタンが強調表示された [アクセス許可] 要求ダイアログをのスクリーンショット](../images/sso-permissions-request.png)

    > [!NOTE]
    > <span data-ttu-id="f3c04-256">ユーザーがこのアクセス許可の要求を受け入れると、今後再びプロンプトが表示されることはありません。</span><span class="sxs-lookup"><span data-stu-id="f3c04-256">After a user accepts this permissions request, they won't be prompted again in the future.</span></span>

7. <span data-ttu-id="f3c04-257">アドインは、サインインしているユーザーの OneDrive for Business からデータを読み取り、上位 10 のファイルとフォルダーの名前を電子メール メッセージの本文に書き込みます。</span><span class="sxs-lookup"><span data-stu-id="f3c04-257">The add-in reads data from the signed-in user's OneDrive for Business and writes the names of the top 10 files and folders to the body of the email message.</span></span>

    ![Outlook のメッセージ作成ウィンドウに OneDrive for Business の情報を示すスクリーンショット](../images/sso-onedrive-info-outlook.png)

## <a name="next-steps"></a><span data-ttu-id="f3c04-259">次の手順</span><span class="sxs-lookup"><span data-stu-id="f3c04-259">Next steps</span></span>

<span data-ttu-id="f3c04-260">これで完了です。SSO クイック スタートで Yeoman ジェネレーターを使用して作成した SSO 対応アドインの機能が [正常にカスタマイズされました](sso-quickstart.md)。</span><span class="sxs-lookup"><span data-stu-id="f3c04-260">Congratulations, you've successfully customized the functionality of the SSO-enabled add-in that you created with the Yeoman generator in the [SSO quick start](sso-quickstart.md).</span></span> <span data-ttu-id="f3c04-261">Yeoman ジェネレーターが自動的に完了した SSO の構成手順、および SSO プロセスを容易にするコードの詳細については、「[シングル サインオンを使用する Node.js Office アドインを作成する](../develop/create-sso-office-add-ins-nodejs.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f3c04-261">To learn more about SSO configuration steps that the Yeoman generator completed automatically, and the code that facilitates the SSO process, see the [Create a Node.js Office Add-in that uses single sign-on](../develop/create-sso-office-add-ins-nodejs.md) tutorial.</span></span>

## <a name="see-also"></a><span data-ttu-id="f3c04-262">関連項目</span><span class="sxs-lookup"><span data-stu-id="f3c04-262">See also</span></span>

- [<span data-ttu-id="f3c04-263">Office アドインのシングル サインオンを有効化する</span><span class="sxs-lookup"><span data-stu-id="f3c04-263">Enable single sign-on for Office Add-ins</span></span>](../develop/sso-in-office-add-ins.md)
- [<span data-ttu-id="f3c04-264">シングル サインオン (SSO) のクイック スタート</span><span class="sxs-lookup"><span data-stu-id="f3c04-264">Single sign-on (SSO) quick start</span></span>](sso-quickstart.md)
- [<span data-ttu-id="f3c04-265">シングル サインオンを使用する Node.js Office アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="f3c04-265">Create a Node.js Office Add-in that uses single sign-on</span></span>](../develop/create-sso-office-add-ins-nodejs.md)
- [<span data-ttu-id="f3c04-266">シングル サインオン (SSO) のエラー メッセージのトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="f3c04-266">Troubleshoot error messages for single sign-on (SSO)</span></span>](../develop/troubleshoot-sso-in-office-add-ins.md)
