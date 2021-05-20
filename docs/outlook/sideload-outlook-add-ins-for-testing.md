---
title: テスト用に Outlook アドインをサイドロードする
description: サイドロードを使用して、最初にアドイン カタログに置かずに、テスト用に Outlook アドインをインストールします。
ms.date: 05/13/2021
localization_priority: Normal
ms.openlocfilehash: 9d0fb246f6522c745658a09fce6934ee44d5079a
ms.sourcegitcommit: 693d364616b42eea66977eef47530adabc51a40f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/19/2021
ms.locfileid: "52555193"
---
# <a name="sideload-outlook-add-ins-for-testing"></a><span data-ttu-id="c813e-103">テスト用に Outlook アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="c813e-103">Sideload Outlook add-ins for testing</span></span>

<span data-ttu-id="c813e-104">サイドロードを使用すると、最初にアドイン カタログに置かなくても、テスト用に Outlook アドインをインストールすることができます。</span><span class="sxs-lookup"><span data-stu-id="c813e-104">You can use sideloading to install an Outlook add-in for testing without having to first put it in an add-in catalog.</span></span>

## <a name="sideload-automatically"></a><span data-ttu-id="c813e-105">サイドロードは自動的に行われます</span><span class="sxs-lookup"><span data-stu-id="c813e-105">Sideload automatically</span></span>

<span data-ttu-id="c813e-106">Office アドイン[用の Yeoman ジェネレーター](https://github.com/OfficeDev/generator-office)を使用して Outlook アドインを作成した場合は、コマンド ラインを使用してサイドローディングを実行することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="c813e-106">If you created your Outlook add-in using [the Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office), sideloading is best done through the command line.</span></span> <span data-ttu-id="c813e-107">これは、1つのコマンドでサポートされているすべてのデバイス全体で私たちのツールとサイドロードを利用します。</span><span class="sxs-lookup"><span data-stu-id="c813e-107">This will take advantage of our tooling and sideload across all of your supported devices in one command.</span></span>

1. <span data-ttu-id="c813e-108">コマンド ラインを使用して、Yeoman によって生成されたアドイン プロジェクトのルート ディレクトリに移動します。</span><span class="sxs-lookup"><span data-stu-id="c813e-108">Using the command line, navigate to the root directory of your Yeoman generated add-in project.</span></span> <span data-ttu-id="c813e-109">コマンド`npm start`を実行します。</span><span class="sxs-lookup"><span data-stu-id="c813e-109">Run the command `npm start`.</span></span>

1. <span data-ttu-id="c813e-110">Outlook アドインは、デスクトップ コンピュータ上のOutlookに自動的にサイドロードされます。</span><span class="sxs-lookup"><span data-stu-id="c813e-110">Your Outlook add-in will automatically sideload to Outlook on your desktop computer.</span></span> <span data-ttu-id="c813e-111">アドインのサイドロードが試行され、マニフェスト ファイルの名前と場所が一覧表示されたダイアログが表示されます。</span><span class="sxs-lookup"><span data-stu-id="c813e-111">You'll see a dialog appear, stating there is an attempt to sideload the add-in, listing the name and the location of the manifest file.</span></span> <span data-ttu-id="c813e-112">**[OK]** を選択すると、マニフェストが登録されます。</span><span class="sxs-lookup"><span data-stu-id="c813e-112">Select **OK**, which will register the manifest.</span></span>

    > [!IMPORTANT]
    > <span data-ttu-id="c813e-113">マニフェストにエラーが含まれている場合、またはマニフェストへのパスが無効な場合は、エラー メッセージが表示されます。</span><span class="sxs-lookup"><span data-stu-id="c813e-113">If the manifest contains an error or the path to the manifest is invalid, you'll receive an error message.</span></span>

1. <span data-ttu-id="c813e-114">マニフェストにエラーがなく、パスが有効な場合、アドインはサイドロードされ、デスクトップと Web 上のOutlookの両方で使用できるようになります。</span><span class="sxs-lookup"><span data-stu-id="c813e-114">If your manifest contains no errors and the path is valid, your add-in will now be sideloaded and available on both your desktop and in Outlook on the web.</span></span> <span data-ttu-id="c813e-115">また、サポートされているすべてのデバイスにインストールされます。</span><span class="sxs-lookup"><span data-stu-id="c813e-115">It will also be installed across all your supported devices.</span></span>

## <a name="sideload-manually"></a><span data-ttu-id="c813e-116">手動でサイドロード</span><span class="sxs-lookup"><span data-stu-id="c813e-116">Sideload manually</span></span>

<span data-ttu-id="c813e-117">前のセクションで説明したように、コマンド ラインを使用して自動的にサイドローディングを行うことを強くお勧めしますが、Outlook クライアントに基づいてOutlookアドインを手動でサイドロードすることもできます。</span><span class="sxs-lookup"><span data-stu-id="c813e-117">Though we strongly recommend sideloading automatically through the command line as covered in the previous section, you can also manually sideload an Outlook add-in based on the Outlook client.</span></span>

### <a name="outlook-on-the-web"></a><span data-ttu-id="c813e-118">Outlook on the web</span><span class="sxs-lookup"><span data-stu-id="c813e-118">Outlook on the web</span></span>

<span data-ttu-id="c813e-119">Web 上のアドインをサイドロードOutlookプロセスは、新しいバージョンとクラシック バージョンのどちらを使用しているかによって異なります。</span><span class="sxs-lookup"><span data-stu-id="c813e-119">The process for sideloading an add-in in Outlook on the web depends upon whether you are using the new or classic version.</span></span>

- <span data-ttu-id="c813e-120">メールボックスのツールバーが次の図のような場合、「[新しい Outlook on the web のアドインをサイドロードする](#new-outlook-on-the-web)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c813e-120">If your mailbox toolbar looks like the following image, see [Sideload an add-in in the new Outlook on the web](#new-outlook-on-the-web).</span></span>

    ![新しい Outlook on the web の部分的なスクリーンショット](../images/outlook-on-the-web-new-toolbar.png)

- <span data-ttu-id="c813e-122">メールボックスのツールバーが次の図のような場合、「[従来の Outlook on the web のアドインをサイドロードする](#classic-outlook-on-the-web)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c813e-122">If your mailbox toolbar looks like the following image, see [Sideload an add-in in classic Outlook on the web](#classic-outlook-on-the-web).</span></span>

    ![従来の Outlook on the web の部分的なスクリーンショット](../images/outlook-on-the-web-classic-toolbar.png)

> [!NOTE]
> <span data-ttu-id="c813e-124">組織のメールボックスのツールバーにロゴが含まれている場合、上の図に示されるものと表示が少し異なる場合があります。</span><span class="sxs-lookup"><span data-stu-id="c813e-124">If your organization has included its logo in the mailbox toolbar, you might see something slightly different than shown in the preceding images.</span></span>

### <a name="new-outlook-on-the-web"></a><span data-ttu-id="c813e-125">Web 上の新しいOutlook</span><span class="sxs-lookup"><span data-stu-id="c813e-125">New Outlook on the web</span></span>

1. <span data-ttu-id="c813e-126">[[Outlook on the web]](https://outlook.office.com) に進みます。</span><span class="sxs-lookup"><span data-stu-id="c813e-126">Go to [Outlook on the web](https://outlook.office.com).</span></span>

1. <span data-ttu-id="c813e-127">新しいメッセージを作成します。</span><span class="sxs-lookup"><span data-stu-id="c813e-127">Create a new message.</span></span>

1. <span data-ttu-id="c813e-128">新しいメッセージの下部で [**...**] を選択し、表示されるメニューから [**アドインを取得**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="c813e-128">Choose **...** from the bottom of the new message and then select **Get Add-ins** from the menu that appears.</span></span>

    ![[アドインを取得] オプションが強調表示された Outlook on the web のメッセージ作成ウィンドウ](../images/outlook-on-the-web-new-get-add-ins.png)

1. <span data-ttu-id="c813e-130">[**Outlook のアドイン**] ダイアログ ボックスで、[**個人用アドイン**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="c813e-130">In the **Add-Ins for Outlook** dialog box, select **My add-ins**.</span></span>

    ![[個人用アドイン] が選択された 新しい Outlook on the web の [Outlook のアドイン] ダイアログ ボックス](../images/outlook-on-the-web-new-my-add-ins.png)

1. <span data-ttu-id="c813e-132">ダイアログ ボックスの下部にある [**カスタム アドイン**] セクションに移動します。</span><span class="sxs-lookup"><span data-stu-id="c813e-132">Locate the **Custom add-ins** section at the bottom of the dialog box.</span></span> <span data-ttu-id="c813e-133">[**カスタム アドインを追加**] リンクを選択し、[**ファイルから追加**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="c813e-133">Select the **Add a custom add-in** link, and then select **Add from file**.</span></span>

    ![ファイル オプションからの追加を示すアドイン スクリーンショットの管理](../images/outlook-sideload-desktop-add-from-file.png)

1. <span data-ttu-id="c813e-p106">カスタム アドインのマニフェスト ファイルを探してインストールします。インストール中にすべてのプロンプトを受け入れます。</span><span class="sxs-lookup"><span data-stu-id="c813e-p106">Locate the manifest file for your custom add-in and install it. Accept all prompts during the installation.</span></span>

### <a name="classic-outlook-on-the-web"></a><span data-ttu-id="c813e-137">ウェブ上の古典的なOutlook</span><span class="sxs-lookup"><span data-stu-id="c813e-137">Classic Outlook on the web</span></span>

1. <span data-ttu-id="c813e-138">[[Outlook on the web]](https://outlook.office.com) に進みます。</span><span class="sxs-lookup"><span data-stu-id="c813e-138">Go to [Outlook on the web](https://outlook.office.com).</span></span>

1. <span data-ttu-id="c813e-139">ツールバー右上のセクションにあるギア アイコンを選択し、[**アドインの管理**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="c813e-139">Choose the gear icon in the top-right section of the toolbar and select **Manage add-ins**.</span></span>

    ![[アドインの管理] オプションを示す Outlook on the web のスクリーンショット](../images/outlook-sideload-web-manage-integrations.png)

1. <span data-ttu-id="c813e-141">**アドインの管理** ページで、**[アドイン]** を選択してから、**[個人用アドイン]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="c813e-141">On the **Manage add-ins** page, select **Add-Ins**, and then select **My add-ins**.</span></span>

    ![Outlook on the web の [ストア] ダイアログ ボックスで [個人用アドイン] を選択しているところ](../images/outlook-sideload-store-select-add-ins.png)

1. <span data-ttu-id="c813e-143">ダイアログ ボックスの下部にある [**カスタム アドイン**] セクションに移動します。</span><span class="sxs-lookup"><span data-stu-id="c813e-143">Locate the **Custom add-ins** section at the bottom of the dialog box.</span></span> <span data-ttu-id="c813e-144">[**カスタム アドインを追加**] リンクを選択し、[**ファイルから追加**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="c813e-144">Select the **Add a custom add-in** link, and then select **Add from file**.</span></span>

    ![ファイル オプションからの追加を示すアドイン スクリーンショットの管理](../images/outlook-sideload-desktop-add-from-file.png)

1. <span data-ttu-id="c813e-p108">カスタム アドインのマニフェスト ファイルを探してインストールします。インストール中にすべてのプロンプトを受け入れます。</span><span class="sxs-lookup"><span data-stu-id="c813e-p108">Locate the manifest file for your custom add-in and install it. Accept all prompts during the installation.</span></span>

### <a name="outlook-on-the-desktop"></a><span data-ttu-id="c813e-148">デスクトップ上のOutlook</span><span class="sxs-lookup"><span data-stu-id="c813e-148">Outlook on the desktop</span></span>

#### <a name="outlook-2016-or-later"></a><span data-ttu-id="c813e-149">Outlook 2016以降</span><span class="sxs-lookup"><span data-stu-id="c813e-149">Outlook 2016 or later</span></span>

1. <span data-ttu-id="c813e-150">Windowsまたは Mac でOutlook 2016以降で開きます。</span><span class="sxs-lookup"><span data-stu-id="c813e-150">Open Outlook 2016 or later on Windows or Mac.</span></span>

1. <span data-ttu-id="c813e-151">リボンで [**アドインを取得**] ボタンを選択します。</span><span class="sxs-lookup"><span data-stu-id="c813e-151">Select the **Get Add-ins** button on the ribbon.</span></span>

    ![[アドインの取得] ボタンをポイントするリボンをOutlook 2016します。](../images/outlook-sideload-desktop-store.png)

    > [!IMPORTANT]
    > <span data-ttu-id="c813e-153">使用しているバージョンのOutlookに [**アドインの取得**] ボタンが表示されない場合は、次のオプションを選択します。</span><span class="sxs-lookup"><span data-stu-id="c813e-153">If you don't see the **Get Add-ins** button in your version of Outlook, select:</span></span>
    >
    > - <span data-ttu-id="c813e-154">リボン上の **[ストア**] ボタン (可能な場合)</span><span class="sxs-lookup"><span data-stu-id="c813e-154">**Store** button on the ribbon, if available.</span></span>
    >
    >   <span data-ttu-id="c813e-155">OR</span><span class="sxs-lookup"><span data-stu-id="c813e-155">OR</span></span>
    >
    > - <span data-ttu-id="c813e-156">**[ファイル]** メニューをクリックし、[**情報**] タブの [**アドインの管理**] ボタンを選択して、Web のOutlookで [**アドイン**] ダイアログを開きます。</span><span class="sxs-lookup"><span data-stu-id="c813e-156">**File** menu, then select the **Manage Add-ins** button on the **Info** tab to open the **Add-ins** dialog in Outlook on the web.</span></span><br><span data-ttu-id="c813e-157">Web エクスペリエンスの詳細については、前のセクション「web 上の[Outlookアドインをサイドロード](#outlook-on-the-web)する」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c813e-157">You can see more about the web experience in the previous section [Sideload an add-in in Outlook on the web](#outlook-on-the-web).</span></span>

1. <span data-ttu-id="c813e-158">ダイアログの上部にタブがある場合は、[ **アドイン** ] タブが選択されていることを確認します。</span><span class="sxs-lookup"><span data-stu-id="c813e-158">If there are tabs near the top of the dialog, ensure that the **Add-ins** tab is selected.</span></span> <span data-ttu-id="c813e-159">[ **個人用アドイン**] を選びます。</span><span class="sxs-lookup"><span data-stu-id="c813e-159">Choose **My add-ins**.</span></span>

    ![Outlook 2016 の [ストア] ダイアログ ボックスで [個人用アドイン] を選択しているところ](../images/outlook-sideload-store-select-add-ins.png)

1. <span data-ttu-id="c813e-161">ダイアログ ボックスの下部にある **[カスタム アドイン]** セクションに移動します。</span><span class="sxs-lookup"><span data-stu-id="c813e-161">Locate the **Custom add-ins** section at the bottom of the dialog.</span></span> <span data-ttu-id="c813e-162">**[カスタム アドインを追加]** リンクを選択し、**[ファイルから追加]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="c813e-162">Select the **Add a custom add-in** link, and then select **Add from file**.</span></span>

    ![[ファイルから追加] オプションを示す [ストア] のスクリーンショット](../images/outlook-sideload-desktop-add-from-file.png)

1. <span data-ttu-id="c813e-p111">カスタム アドインのマニフェスト ファイルを探してインストールします。インストール中にすべてのプロンプトを受け入れます。</span><span class="sxs-lookup"><span data-stu-id="c813e-p111">Locate the manifest file for your custom add-in and install it. Accept all prompts during the installation.</span></span>

#### <a name="outlook-2013"></a><span data-ttu-id="c813e-166">Outlook 2013</span><span class="sxs-lookup"><span data-stu-id="c813e-166">Outlook 2013</span></span>

1. <span data-ttu-id="c813e-167">Windows Outlook 2013をオープンします。</span><span class="sxs-lookup"><span data-stu-id="c813e-167">Open Outlook 2013 on Windows.</span></span>

1. <span data-ttu-id="c813e-168">[**ファイル]** メニューを選択し、[**情報**] タブ **の [アドインの管理**] ボタンを選択Outlook、ブラウザで Web バージョンを開きます。</span><span class="sxs-lookup"><span data-stu-id="c813e-168">Select the **File** menu, then select the **Manage Add-ins** button on the **Info** tab. Outlook will open the web version in a browser.</span></span>

1. <span data-ttu-id="c813e-169">Web 上のOutlookのバージョン[に従って、web セクションの [Outlookのアドインをサイドロード](#outlook-on-the-web)するの手順に従います。</span><span class="sxs-lookup"><span data-stu-id="c813e-169">Follow the steps in the [Sideload an add-in in Outlook on the web](#outlook-on-the-web) section according to your version of Outlook on the web.</span></span>

## <a name="remove-a-sideloaded-add-in"></a><span data-ttu-id="c813e-170">サイドローデッド アドインを削除する</span><span class="sxs-lookup"><span data-stu-id="c813e-170">Remove a sideloaded add-in</span></span>

<span data-ttu-id="c813e-171">Outlookのすべてのバージョンで、サイドローデッド アドインを削除するキーは、インストールされている **アドイン** を一覧表示する [マイ アドイン] ダイアログです。アドインの省略記号 ( `...` ) を選択し、[**削除**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="c813e-171">On all versions of Outlook, the key to removing a sideloaded add-in is the **My Add-ins** dialog which lists your installed add-ins. Choose the ellipsis (`...`) for the add-in then select **Remove**.</span></span>

<span data-ttu-id="c813e-172">Outlook クライアントの [**個人用アドイン**] ダイアログ ボックスに移動するには、この記事の前のセクションで [説明した手動サイドローディング](#sideload-manually)の最後の手順を使用します。</span><span class="sxs-lookup"><span data-stu-id="c813e-172">To navigate to the **My Add-ins** dialog box for your Outlook client, use the last steps listed for [manual sideloading](#sideload-manually) in the previous sections of this article.</span></span>

<span data-ttu-id="c813e-173">サイドローデッド アドインをOutlookから削除するには、この記事で説明した手順を使用して、インストールされているアドインの一覧が表示されるダイアログ ボックスの **[カスタム アドイン**] セクションでアドインを検索します。アドインの省略記号 ( `...` ) を選択し、[**削除**] を選択して、特定のアドインを削除します。</span><span class="sxs-lookup"><span data-stu-id="c813e-173">To remove a sideloaded add-in from Outlook, use the steps previously described in this article to find the add-in in the **Custom add-ins** section of the dialog box that lists your installed add-ins. Choose the ellipsis (`...`) for the add-in then choose **Remove** to remove that specific add-in.</span></span> <span data-ttu-id="c813e-174">ダイアログを閉じます。</span><span class="sxs-lookup"><span data-stu-id="c813e-174">Close the dialog.</span></span>
