---
title: テスト用に Outlook アドインをサイドロードする
description: サイドロードを使用して、最初にアドイン カタログに置かずに、テスト用に Outlook アドインをインストールします。
ms.date: 05/13/2021
localization_priority: Normal
ms.openlocfilehash: 47eb5da19f858b6e30339acc59da24a818fc0959
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/23/2021
ms.locfileid: "53077030"
---
# <a name="sideload-outlook-add-ins-for-testing"></a><span data-ttu-id="7f5ce-103">テスト用に Outlook アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="7f5ce-103">Sideload Outlook add-ins for testing</span></span>

<span data-ttu-id="7f5ce-104">サイドロードを使用すると、最初にアドイン カタログに置かなくても、テスト用に Outlook アドインをインストールすることができます。</span><span class="sxs-lookup"><span data-stu-id="7f5ce-104">You can use sideloading to install an Outlook add-in for testing without having to first put it in an add-in catalog.</span></span>

## <a name="sideload-automatically"></a><span data-ttu-id="7f5ce-105">サイドロードが自動的に実行される</span><span class="sxs-lookup"><span data-stu-id="7f5ce-105">Sideload automatically</span></span>

<span data-ttu-id="7f5ce-106">Office アドイン用の[Yeoman](https://github.com/OfficeDev/generator-office)ジェネレーターを使用して Outlook アドインを作成した場合は、コマンド ラインを使用してサイドローディングを行うのが最適です。</span><span class="sxs-lookup"><span data-stu-id="7f5ce-106">If you created your Outlook add-in using [the Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office), sideloading is best done through the command line.</span></span> <span data-ttu-id="7f5ce-107">これにより、1 つのコマンドでサポートされているすべてのデバイスでツールとサイドロードを利用できます。</span><span class="sxs-lookup"><span data-stu-id="7f5ce-107">This will take advantage of our tooling and sideload across all of your supported devices in one command.</span></span>

1. <span data-ttu-id="7f5ce-108">コマンド ラインを使用して、Yeoman が生成したアドイン プロジェクトのルート ディレクトリに移動します。</span><span class="sxs-lookup"><span data-stu-id="7f5ce-108">Using the command line, navigate to the root directory of your Yeoman generated add-in project.</span></span> <span data-ttu-id="7f5ce-109">コマンド`npm start`を実行します。</span><span class="sxs-lookup"><span data-stu-id="7f5ce-109">Run the command `npm start`.</span></span>

1. <span data-ttu-id="7f5ce-110">ユーザー Outlookは、デスクトップ コンピューター上のOutlookに自動的にサイドロードされます。</span><span class="sxs-lookup"><span data-stu-id="7f5ce-110">Your Outlook add-in will automatically sideload to Outlook on your desktop computer.</span></span> <span data-ttu-id="7f5ce-111">アドインをサイドロードしようとして、マニフェスト ファイルの名前と場所を一覧に表示するダイアログが表示されます。</span><span class="sxs-lookup"><span data-stu-id="7f5ce-111">You'll see a dialog appear, stating there is an attempt to sideload the add-in, listing the name and the location of the manifest file.</span></span> <span data-ttu-id="7f5ce-112">**[OK]** を選択し、マニフェストを登録します。</span><span class="sxs-lookup"><span data-stu-id="7f5ce-112">Select **OK**, which will register the manifest.</span></span>

    > [!IMPORTANT]
    > <span data-ttu-id="7f5ce-113">マニフェストにエラーが含まれているか、マニフェストへのパスが無効な場合は、エラー メッセージが表示されます。</span><span class="sxs-lookup"><span data-stu-id="7f5ce-113">If the manifest contains an error or the path to the manifest is invalid, you'll receive an error message.</span></span>

1. <span data-ttu-id="7f5ce-114">マニフェストにエラーが含まれるが、パスが有効な場合、アドインはサイドロードされ、デスクトップと Outlook on the web の両方で使用できます。</span><span class="sxs-lookup"><span data-stu-id="7f5ce-114">If your manifest contains no errors and the path is valid, your add-in will now be sideloaded and available on both your desktop and in Outlook on the web.</span></span> <span data-ttu-id="7f5ce-115">また、サポートされているすべてのデバイスにインストールされます。</span><span class="sxs-lookup"><span data-stu-id="7f5ce-115">It will also be installed across all your supported devices.</span></span>

## <a name="sideload-manually"></a><span data-ttu-id="7f5ce-116">サイドロードを手動で実行する</span><span class="sxs-lookup"><span data-stu-id="7f5ce-116">Sideload manually</span></span>

<span data-ttu-id="7f5ce-117">前のセクションで説明したコマンド ラインから自動的にサイドローディングすることを強く推奨しますが、Outlook クライアントに基づいて Outlook アドインを手動でサイドロードすることもできます。</span><span class="sxs-lookup"><span data-stu-id="7f5ce-117">Though we strongly recommend sideloading automatically through the command line as covered in the previous section, you can also manually sideload an Outlook add-in based on the Outlook client.</span></span>

### <a name="outlook-on-the-web"></a><span data-ttu-id="7f5ce-118">Outlook on the web</span><span class="sxs-lookup"><span data-stu-id="7f5ce-118">Outlook on the web</span></span>

<span data-ttu-id="7f5ce-119">新しいバージョンまたはクラシック バージョンを使用Outlook on the webアドインをサイドローディングするプロセスは異なります。</span><span class="sxs-lookup"><span data-stu-id="7f5ce-119">The process for sideloading an add-in in Outlook on the web depends upon whether you are using the new or classic version.</span></span>

- <span data-ttu-id="7f5ce-120">メールボックスのツールバーが次の図のような場合、「[新しい Outlook on the web のアドインをサイドロードする](#new-outlook-on-the-web)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="7f5ce-120">If your mailbox toolbar looks like the following image, see [Sideload an add-in in the new Outlook on the web](#new-outlook-on-the-web).</span></span>

    ![新しいツール バーの部分的なOutlook on the webです。](../images/outlook-on-the-web-new-toolbar.png)

- <span data-ttu-id="7f5ce-122">メールボックスのツールバーが次の図のような場合、「[従来の Outlook on the web のアドインをサイドロードする](#classic-outlook-on-the-web)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="7f5ce-122">If your mailbox toolbar looks like the following image, see [Sideload an add-in in classic Outlook on the web](#classic-outlook-on-the-web).</span></span>

    ![従来のツール バーの一部Outlook on the webスクリーンショット。](../images/outlook-on-the-web-classic-toolbar.png)

> [!NOTE]
> <span data-ttu-id="7f5ce-124">組織のメールボックスのツールバーにロゴが含まれている場合、上の図に示されるものと表示が少し異なる場合があります。</span><span class="sxs-lookup"><span data-stu-id="7f5ce-124">If your organization has included its logo in the mailbox toolbar, you might see something slightly different than shown in the preceding images.</span></span>

### <a name="new-outlook-on-the-web"></a><span data-ttu-id="7f5ce-125">新しいOutlook on the web</span><span class="sxs-lookup"><span data-stu-id="7f5ce-125">New Outlook on the web</span></span>

1. <span data-ttu-id="7f5ce-126">[[Outlook on the web]](https://outlook.office.com) に進みます。</span><span class="sxs-lookup"><span data-stu-id="7f5ce-126">Go to [Outlook on the web](https://outlook.office.com).</span></span>

1. <span data-ttu-id="7f5ce-127">新しいメッセージを作成します。</span><span class="sxs-lookup"><span data-stu-id="7f5ce-127">Create a new message.</span></span>

1. <span data-ttu-id="7f5ce-128">新しいメッセージの下部で [**...**] を選択し、表示されるメニューから [**アドインを取得**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="7f5ce-128">Choose **...** from the bottom of the new message and then select **Get Add-ins** from the menu that appears.</span></span>

    ![[アドインの取得] オプションが強調表示Outlook on the web新しいウィンドウのメッセージ作成ウィンドウ。](../images/outlook-on-the-web-new-get-add-ins.png)

1. <span data-ttu-id="7f5ce-130">[**Outlook のアドイン**] ダイアログ ボックスで、[**個人用アドイン**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="7f5ce-130">In the **Add-Ins for Outlook** dialog box, select **My add-ins**.</span></span>

    ![[自分のOutlook] を選択した新しいOutlook on the webダイアログ ボックスのアドイン。](../images/outlook-on-the-web-new-my-add-ins.png)

1. <span data-ttu-id="7f5ce-132">ダイアログ ボックスの下部にある [**カスタム アドイン**] セクションに移動します。</span><span class="sxs-lookup"><span data-stu-id="7f5ce-132">Locate the **Custom add-ins** section at the bottom of the dialog box.</span></span> <span data-ttu-id="7f5ce-133">[**カスタム アドインを追加**] リンクを選択し、[**ファイルから追加**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="7f5ce-133">Select the **Add a custom add-in** link, and then select **Add from file**.</span></span>

    ![[ファイルから追加] オプションをポイントするアドインのスクリーンショットを管理します。](../images/outlook-sideload-desktop-add-from-file.png)

1. <span data-ttu-id="7f5ce-p106">カスタム アドインのマニフェスト ファイルを探してインストールします。インストール中にすべてのプロンプトを受け入れます。</span><span class="sxs-lookup"><span data-stu-id="7f5ce-p106">Locate the manifest file for your custom add-in and install it. Accept all prompts during the installation.</span></span>

### <a name="classic-outlook-on-the-web"></a><span data-ttu-id="7f5ce-137">クラシック Outlook on the web</span><span class="sxs-lookup"><span data-stu-id="7f5ce-137">Classic Outlook on the web</span></span>

1. <span data-ttu-id="7f5ce-138">[[Outlook on the web]](https://outlook.office.com) に進みます。</span><span class="sxs-lookup"><span data-stu-id="7f5ce-138">Go to [Outlook on the web](https://outlook.office.com).</span></span>

1. <span data-ttu-id="7f5ce-139">ツールバー右上のセクションにあるギア アイコンを選択し、[**アドインの管理**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="7f5ce-139">Choose the gear icon in the top-right section of the toolbar and select **Manage add-ins**.</span></span>

    ![Outlook on the webアドインの管理オプションをポイントするスクリーンショットを作成します。](../images/outlook-sideload-web-manage-integrations.png)

1. <span data-ttu-id="7f5ce-141">**アドインの管理** ページで、**[アドイン]** を選択してから、**[個人用アドイン]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="7f5ce-141">On the **Manage add-ins** page, select **Add-Ins**, and then select **My add-ins**.</span></span>

    ![Outlook on the web[マイ アドイン] が選択されている場合は、[ストア] ダイアログボックスを開きます。](../images/outlook-sideload-store-select-add-ins.png)

1. <span data-ttu-id="7f5ce-143">ダイアログ ボックスの下部にある [**カスタム アドイン**] セクションに移動します。</span><span class="sxs-lookup"><span data-stu-id="7f5ce-143">Locate the **Custom add-ins** section at the bottom of the dialog box.</span></span> <span data-ttu-id="7f5ce-144">[**カスタム アドインを追加**] リンクを選択し、[**ファイルから追加**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="7f5ce-144">Select the **Add a custom add-in** link, and then select **Add from file**.</span></span>

    ![[ファイルから追加] オプションをポイントするアドインのスクリーンショットを管理します。](../images/outlook-sideload-desktop-add-from-file.png)

1. <span data-ttu-id="7f5ce-p108">カスタム アドインのマニフェスト ファイルを探してインストールします。インストール中にすべてのプロンプトを受け入れます。</span><span class="sxs-lookup"><span data-stu-id="7f5ce-p108">Locate the manifest file for your custom add-in and install it. Accept all prompts during the installation.</span></span>

### <a name="outlook-on-the-desktop"></a><span data-ttu-id="7f5ce-148">Outlookの設定</span><span class="sxs-lookup"><span data-stu-id="7f5ce-148">Outlook on the desktop</span></span>

#### <a name="outlook-2016-or-later"></a><span data-ttu-id="7f5ce-149">Outlook 2016以降</span><span class="sxs-lookup"><span data-stu-id="7f5ce-149">Outlook 2016 or later</span></span>

1. <span data-ttu-id="7f5ce-150">[Outlook 2016または Mac で、Windows以降を開きます。</span><span class="sxs-lookup"><span data-stu-id="7f5ce-150">Open Outlook 2016 or later on Windows or Mac.</span></span>

1. <span data-ttu-id="7f5ce-151">リボンで [**アドインを取得**] ボタンを選択します。</span><span class="sxs-lookup"><span data-stu-id="7f5ce-151">Select the **Get Add-ins** button on the ribbon.</span></span>

    ![Outlook 2016アドインの取得] ボタンをポイントするリボンを選択します。](../images/outlook-sideload-desktop-store.png)

    > [!IMPORTANT]
    > <span data-ttu-id="7f5ce-153">バージョンの [アドインの取得] ボタンが表示されていない場合は、次Outlook選択します。</span><span class="sxs-lookup"><span data-stu-id="7f5ce-153">If you don't see the **Get Add-ins** button in your version of Outlook, select:</span></span>
    >
    > - <span data-ttu-id="7f5ce-154">**リボン** の [ストア] ボタン (使用可能な場合)。</span><span class="sxs-lookup"><span data-stu-id="7f5ce-154">**Store** button on the ribbon, if available.</span></span>
    >
    >   <span data-ttu-id="7f5ce-155">または</span><span class="sxs-lookup"><span data-stu-id="7f5ce-155">OR</span></span>
    >
    > - <span data-ttu-id="7f5ce-156">**[** ファイル] メニューの [情報] タブの[アドインの管理]ボタンを選択して、[アドイン] ダイアログボックスを開Outlook on the web。 </span><span class="sxs-lookup"><span data-stu-id="7f5ce-156">**File** menu, then select the **Manage Add-ins** button on the **Info** tab to open the **Add-ins** dialog in Outlook on the web.</span></span><br><span data-ttu-id="7f5ce-157">Web エクスペリエンスの詳細については、前のセクションの「アドインをサイドロードする」を参照[Outlook on the web。](#outlook-on-the-web)</span><span class="sxs-lookup"><span data-stu-id="7f5ce-157">You can see more about the web experience in the previous section [Sideload an add-in in Outlook on the web](#outlook-on-the-web).</span></span>

1. <span data-ttu-id="7f5ce-158">ダイアログの上部にタブがある場合は、[アドイン] タブが **選択** されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="7f5ce-158">If there are tabs near the top of the dialog, ensure that the **Add-ins** tab is selected.</span></span> <span data-ttu-id="7f5ce-159">[ **個人用アドイン**] を選びます。</span><span class="sxs-lookup"><span data-stu-id="7f5ce-159">Choose **My add-ins**.</span></span>

    ![Outlook 2016[マイ アドイン] が選択されている場合は、[ストア] ダイアログボックスを開きます。](../images/outlook-sideload-store-select-add-ins.png)

1. <span data-ttu-id="7f5ce-161">ダイアログ ボックスの下部にある **[カスタム アドイン]** セクションに移動します。</span><span class="sxs-lookup"><span data-stu-id="7f5ce-161">Locate the **Custom add-ins** section at the bottom of the dialog.</span></span> <span data-ttu-id="7f5ce-162">**[カスタム アドインを追加]** リンクを選択し、**[ファイルから追加]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="7f5ce-162">Select the **Add a custom add-in** link, and then select **Add from file**.</span></span>

    ![[ファイルから追加] オプションをポイントするスクリーンショットを保存します。](../images/outlook-sideload-desktop-add-from-file.png)

1. <span data-ttu-id="7f5ce-p111">カスタム アドインのマニフェスト ファイルを探してインストールします。インストール中にすべてのプロンプトを受け入れます。</span><span class="sxs-lookup"><span data-stu-id="7f5ce-p111">Locate the manifest file for your custom add-in and install it. Accept all prompts during the installation.</span></span>

#### <a name="outlook-2013"></a><span data-ttu-id="7f5ce-166">Outlook 2013</span><span class="sxs-lookup"><span data-stu-id="7f5ce-166">Outlook 2013</span></span>

1. <span data-ttu-id="7f5ce-167">2013 Outlook 2013 を開Windows。</span><span class="sxs-lookup"><span data-stu-id="7f5ce-167">Open Outlook 2013 on Windows.</span></span>

1. <span data-ttu-id="7f5ce-168">[ファイル **] メニューを** 選択し、[情報] タブの [アドインの管理] ボタンを選択しますOutlookブラウザーで Web バージョンが開きます。</span><span class="sxs-lookup"><span data-stu-id="7f5ce-168">Select the **File** menu, then select the **Manage Add-ins** button on the **Info** tab. Outlook will open the web version in a browser.</span></span>

1. <span data-ttu-id="7f5ce-169">[アドインのサイドロード][セクションの](#outlook-on-the-web)手順に従って、Outlook on the webのバージョンに従Outlook on the web。</span><span class="sxs-lookup"><span data-stu-id="7f5ce-169">Follow the steps in the [Sideload an add-in in Outlook on the web](#outlook-on-the-web) section according to your version of Outlook on the web.</span></span>

## <a name="remove-a-sideloaded-add-in"></a><span data-ttu-id="7f5ce-170">サイドロードされたアドインを削除する</span><span class="sxs-lookup"><span data-stu-id="7f5ce-170">Remove a sideloaded add-in</span></span>

<span data-ttu-id="7f5ce-171">すべてのバージョンの Outlook、サイドロードされたアドインを削除するキーは、インストールされているアドインを一覧表示する[マイ アドイン] ダイアログです。アドインの省略記号 ( `...` ) を選択し、[削除] を **選択します**。</span><span class="sxs-lookup"><span data-stu-id="7f5ce-171">On all versions of Outlook, the key to removing a sideloaded add-in is the **My Add-ins** dialog which lists your installed add-ins. Choose the ellipsis (`...`) for the add-in then select **Remove**.</span></span>

<span data-ttu-id="7f5ce-172">Outlook クライアントの[マイ アドイン] ダイアログ ボックスに移動するには、この記事の前のセクションで手動[](#sideload-manually)サイドローディングの最後の手順を使用します。</span><span class="sxs-lookup"><span data-stu-id="7f5ce-172">To navigate to the **My Add-ins** dialog box for your Outlook client, use the last steps listed for [manual sideloading](#sideload-manually) in the previous sections of this article.</span></span>

<span data-ttu-id="7f5ce-173">サイドロードされたアドインを Outlook から削除するには、この記事で説明した手順を使用して、インストールされているアドインを一覧表示するダイアログボックスの [カスタム アドイン] セクションでアドインを検索します。アドインの省略記号 ( ) を選択し、[削除] を選択して、その特定の `...` アドインを削除します。 </span><span class="sxs-lookup"><span data-stu-id="7f5ce-173">To remove a sideloaded add-in from Outlook, use the steps previously described in this article to find the add-in in the **Custom add-ins** section of the dialog box that lists your installed add-ins. Choose the ellipsis (`...`) for the add-in then choose **Remove** to remove that specific add-in.</span></span> <span data-ttu-id="7f5ce-174">ダイアログを閉じます。</span><span class="sxs-lookup"><span data-stu-id="7f5ce-174">Close the dialog.</span></span>
