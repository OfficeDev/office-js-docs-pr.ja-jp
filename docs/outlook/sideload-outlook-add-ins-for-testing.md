---
title: テスト用に Outlook アドインをサイドロードする
description: サイドロードを使用して、最初にアドイン カタログに置かずに、テスト用に Outlook アドインをインストールします。
ms.date: 07/09/2020
localization_priority: Normal
ms.openlocfilehash: 9b44b988ddd6552d5f7d14088a0b6f3ae1e410ed
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093883"
---
# <a name="sideload-outlook-add-ins-for-testing"></a><span data-ttu-id="a6518-103">テスト用に Outlook アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="a6518-103">Sideload Outlook add-ins for testing</span></span>

<span data-ttu-id="a6518-104">サイドロードを使用すると、最初にアドイン カタログに置かなくても、テスト用に Outlook アドインをインストールすることができます。</span><span class="sxs-lookup"><span data-stu-id="a6518-104">You can use sideloading to install an Outlook add-in for testing without having to first put it in an add-in catalog.</span></span>

## <a name="sideload-an-add-in-in-outlook-on-the-web"></a><span data-ttu-id="a6518-105">Outlook on the web のアドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="a6518-105">Sideload an add-in in Outlook on the web</span></span>

<span data-ttu-id="a6518-106">Web 上の Outlook でアドインをサイドロードするためのプロセスは、新しいバージョンとクラシックバージョンのどちらを使用しているかによって異なります。</span><span class="sxs-lookup"><span data-stu-id="a6518-106">The process for sideloading an add-in in Outlook on the web depends upon whether you are using the new or classic version.</span></span>

- <span data-ttu-id="a6518-107">メールボックスのツールバーが次の図のような場合、「[新しい Outlook on the web のアドインをサイドロードする](#sideload-an-add-in-in-the-new-outlook-on-the-web)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="a6518-107">If your mailbox toolbar looks like the following image, see [Sideload an add-in in the new Outlook on the web](#sideload-an-add-in-in-the-new-outlook-on-the-web).</span></span>

    ![新しい Outlook on the web の部分的なスクリーンショット](../images/outlook-on-the-web-new-toolbar.png)

- <span data-ttu-id="a6518-109">メールボックスのツールバーが次の図のような場合、「[従来の Outlook on the web のアドインをサイドロードする](#sideload-an-add-in-in-classic-outlook-on-the-web)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="a6518-109">If your mailbox toolbar looks like the following image, see [Sideload an add-in in classic Outlook on the web](#sideload-an-add-in-in-classic-outlook-on-the-web).</span></span>

    ![従来の Outlook on the web の部分的なスクリーンショット](../images/outlook-on-the-web-classic-toolbar.png)

> [!NOTE]
> <span data-ttu-id="a6518-111">組織のメールボックスのツールバーにロゴが含まれている場合、上の図に示されるものと表示が少し異なる場合があります。</span><span class="sxs-lookup"><span data-stu-id="a6518-111">If your organization has included its logo in the mailbox toolbar, you might see something slightly different than shown in the preceding images.</span></span>

### <a name="sideload-an-add-in-in-the-new-outlook-on-the-web"></a><span data-ttu-id="a6518-112">新しい Outlook on the web のアドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="a6518-112">Sideload an add-in in the new Outlook on the web</span></span>

1. <span data-ttu-id="a6518-113">[Office 365 の Outlook](https://outlook.office.com) に移動します。</span><span class="sxs-lookup"><span data-stu-id="a6518-113">Go to [Outlook in Office 365](https://outlook.office.com).</span></span>

1. <span data-ttu-id="a6518-114">Outlook on the web で新しいメッセージを作成します。</span><span class="sxs-lookup"><span data-stu-id="a6518-114">In Outlook on the web, create a new message.</span></span>

1. <span data-ttu-id="a6518-115">新しいメッセージの下部で [**...**] を選択し、表示されるメニューから [**アドインを取得**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="a6518-115">Choose **...** from the bottom of the new message and then select **Get Add-ins** from the menu that appears.</span></span>

    ![[アドインを取得] オプションが強調表示された Outlook on the web のメッセージ作成ウィンドウ](../images/outlook-on-the-web-new-get-add-ins.png)

1. <span data-ttu-id="a6518-117">[**Outlook のアドイン**] ダイアログ ボックスで、[**個人用アドイン**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="a6518-117">In the **Add-Ins for Outlook** dialog box, select **My add-ins**.</span></span>

    ![[個人用アドイン] が選択された 新しい Outlook on the web の [Outlook のアドイン] ダイアログ ボックス](../images/outlook-on-the-web-new-my-add-ins.png)

1. <span data-ttu-id="a6518-119">ダイアログ ボックスの下部にある [**カスタム アドイン**] セクションに移動します。</span><span class="sxs-lookup"><span data-stu-id="a6518-119">Locate the **Custom add-ins** section at the bottom of the dialog box.</span></span> <span data-ttu-id="a6518-120">[**カスタム アドインを追加**] リンクを選択し、[**ファイルから追加**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="a6518-120">Select the **Add a custom add-in** link, and then select **Add from file**.</span></span>

    ![ファイル オプションからの追加を示すアドイン スクリーンショットの管理](../images/outlook-sideload-desktop-add-from-file.png)

1. <span data-ttu-id="a6518-p102">カスタム アドインのマニフェスト ファイルを探してインストールします。インストール中にすべてのプロンプトを受け入れます。</span><span class="sxs-lookup"><span data-stu-id="a6518-p102">Locate the manifest file for your custom add-in and install it. Accept all prompts during the installation.</span></span>

### <a name="sideload-an-add-in-in-classic-outlook-on-the-web"></a><span data-ttu-id="a6518-124">従来の Outlook on the web のアドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="a6518-124">Sideload an add-in in classic Outlook on the web</span></span>

1. <span data-ttu-id="a6518-125">[Office 365 の Outlook](https://outlook.office.com) に移動します。</span><span class="sxs-lookup"><span data-stu-id="a6518-125">Go to [Outlook in Office 365](https://outlook.office.com).</span></span>

1. <span data-ttu-id="a6518-126">ツールバー右上のセクションにあるギア アイコンを選択し、[**アドインの管理**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="a6518-126">Choose the gear icon in the top-right section of the toolbar and select **Manage add-ins**.</span></span>

    ![[アドインの管理] オプションを示す Outlook on the web のスクリーンショット](../images/outlook-sideload-web-manage-integrations.png)

1. <span data-ttu-id="a6518-128">**アドインの管理**ページで、**[アドイン]** を選択してから、**[個人用アドイン]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="a6518-128">On the **Manage add-ins** page, select **Add-Ins**, and then select **My add-ins**.</span></span>

    ![Outlook on the web の [ストア] ダイアログ ボックスで [個人用アドイン] を選択しているところ](../images/outlook-sideload-store-select-add-ins.png)

1. <span data-ttu-id="a6518-130">ダイアログ ボックスの下部にある [**カスタム アドイン**] セクションに移動します。</span><span class="sxs-lookup"><span data-stu-id="a6518-130">Locate the **Custom add-ins** section at the bottom of the dialog box.</span></span> <span data-ttu-id="a6518-131">[**カスタム アドインを追加**] リンクを選択し、[**ファイルから追加**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="a6518-131">Select the **Add a custom add-in** link, and then select **Add from file**.</span></span>

    ![ファイル オプションからの追加を示すアドイン スクリーンショットの管理](../images/outlook-sideload-desktop-add-from-file.png)

1. <span data-ttu-id="a6518-p104">カスタム アドインのマニフェスト ファイルを探してインストールします。インストール中にすべてのプロンプトを受け入れます。</span><span class="sxs-lookup"><span data-stu-id="a6518-p104">Locate the manifest file for your custom add-in and install it. Accept all prompts during the installation.</span></span>

## <a name="sideload-an-add-in-in-outlook-on-the-desktop"></a><span data-ttu-id="a6518-135">Outlook on the desktop のアドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="a6518-135">Sideload an add-in in Outlook on the desktop</span></span>

### <a name="outlook-2016-or-later"></a><span data-ttu-id="a6518-136">Outlook 2016 以降</span><span class="sxs-lookup"><span data-stu-id="a6518-136">Outlook 2016 or later</span></span>

1. <span data-ttu-id="a6518-137">Windows または Mac で Outlook 2016 以降を開きます。</span><span class="sxs-lookup"><span data-stu-id="a6518-137">Open Outlook 2016 or later on Windows or Mac.</span></span>

1. <span data-ttu-id="a6518-138">リボンで [**アドインを取得**] ボタンを選択します。</span><span class="sxs-lookup"><span data-stu-id="a6518-138">Select the **Get Add-ins** button on the ribbon.</span></span>

    ![[ストア] ボタンを示す Outlook 2016 リボン](../images/outlook-sideload-desktop-store.png)

    > [!NOTE]
    > <span data-ttu-id="a6518-140">お使いのバージョンの Outlook で [**アドインを取得**] ボタンが表示されない場合、代わりに、リボンで [**ストア**] ボタンを選択します。</span><span class="sxs-lookup"><span data-stu-id="a6518-140">If you don't see the **Get Add-ins** button in your version of Outlook, select the **Store** button on the ribbon instead.</span></span>

1. <span data-ttu-id="a6518-141">[**アドイン**] を選択し、[**個人用アドイン**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="a6518-141">Select **Add-Ins**, and then select **My add-ins**.</span></span>

    ![Outlook 2016 の [ストア] ダイアログ ボックスで [個人用アドイン] を選択しているところ](../images/outlook-sideload-store-select-add-ins.png)

1. <span data-ttu-id="a6518-143">ダイアログ ボックスの下部にある **[カスタム アドイン]** セクションに移動します。</span><span class="sxs-lookup"><span data-stu-id="a6518-143">Locate the **Custom add-ins** section at the bottom of the dialog.</span></span> <span data-ttu-id="a6518-144">**[カスタム アドインを追加]** リンクを選択し、**[ファイルから追加]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="a6518-144">Select the **Add a custom add-in** link, and then select **Add from file**.</span></span>

    ![[ファイルから追加] オプションを示す [ストア] のスクリーンショット](../images/outlook-sideload-desktop-add-from-file.png)

1. <span data-ttu-id="a6518-p106">カスタム アドインのマニフェスト ファイルを探してインストールします。インストール中にすべてのプロンプトを受け入れます。</span><span class="sxs-lookup"><span data-stu-id="a6518-p106">Locate the manifest file for your custom add-in and install it. Accept all prompts during the installation.</span></span>

### <a name="outlook-2013"></a><span data-ttu-id="a6518-148">Outlook 2013</span><span class="sxs-lookup"><span data-stu-id="a6518-148">Outlook 2013</span></span>

1. <span data-ttu-id="a6518-149">Windows で Outlook 2013 を開きます。</span><span class="sxs-lookup"><span data-stu-id="a6518-149">Open Outlook 2013 on Windows.</span></span>

1. <span data-ttu-id="a6518-150">[**ファイル**] メニューを選択し、[**情報**] タブの [アドインの**管理**] をクリックします。 Outlook はブラウザーを開きます。</span><span class="sxs-lookup"><span data-stu-id="a6518-150">Select the **File** menu, then select the **Manage Add-ins** button on the **Info** tab. Outlook will open a browser.</span></span>

1. <span data-ttu-id="a6518-151">Web 上の Outlook のバージョンに応じて、「 [web 上の outlook でアドインをサイドロード](#sideload-an-add-in-in-outlook-on-the-web)する」セクションの手順を実行します。</span><span class="sxs-lookup"><span data-stu-id="a6518-151">Follow the steps in the [Sideload an add-in in Outlook on the web](#sideload-an-add-in-in-outlook-on-the-web) section according to your version of Outlook on the web.</span></span>

## <a name="remove-a-sideloaded-add-in"></a><span data-ttu-id="a6518-152">サイドロードアドインを削除する</span><span class="sxs-lookup"><span data-stu-id="a6518-152">Remove a sideloaded add-in</span></span>

<span data-ttu-id="a6518-153">サイドロードアドインを Outlook から削除するには、この記事で前述した手順を使用して、インストールされているアドインを一覧表示するダイアログボックスの [**カスタムアドイン**] セクションでアドインを検索します。アドインの省略記号 () を選択 `...` し、[**削除**] を選択して、その特定のアドインを削除します。</span><span class="sxs-lookup"><span data-stu-id="a6518-153">To remove a sideloaded add-in from Outlook, use the steps previously described in this article to find the add-in in the **Custom add-ins** section of the dialog box that lists your installed add-ins. Choose the ellipsis (`...`) for the the add-in and then choose **Remove** to remove that specific add-in.</span></span>