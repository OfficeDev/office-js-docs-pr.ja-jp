---
title: Office の最新バージョンをインストールする
description: Office の最新のビルドを取得するを有効にする方法に関する情報です。
ms.date: 12/04/2017
ms.openlocfilehash: 14e26d9fa9f7ec3b2724cbf2e9787cde9dbe4094
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2018
ms.locfileid: "23943881"
---
# <a name="install-the-latest-version-of-office"></a><span data-ttu-id="1025b-103">Office の最新バージョンをインストールする</span><span class="sxs-lookup"><span data-stu-id="1025b-103">Install the latest version of Office</span></span>

<span data-ttu-id="1025b-104">新しい開発者用機能 (現時点ではプレビュー版のものを含む) は、Office の最新ビルドの取得をオプトインしたサブスクライバーに最初に配信されます。</span><span class="sxs-lookup"><span data-stu-id="1025b-104">New developer features, including those still in preview, are delivered first to subscribers who opt in to get the latest builds of Office.</span></span> 

## <a name="opt-in-to-getting-the-latest-builds"></a><span data-ttu-id="1025b-105">最新のビルドを取得するためにオプトインする</span><span class="sxs-lookup"><span data-stu-id="1025b-105">Opt in to getting the latest builds</span></span>

<span data-ttu-id="1025b-106">Office 2016 の最新ビルドを取得するためにオプトインするには:</span><span class="sxs-lookup"><span data-stu-id="1025b-106">To opt in to getting the latest builds of Office 2016:</span></span> 

- <span data-ttu-id="1025b-107">Office 365 Home、Personal、または University のサブスクライバーは、「[Office Insider プログラム](https://products.office.com/office-insider)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1025b-107">If you're an Office 365 Home, Personal, or University subscriber, see [Be an Office Insider](https://products.office.com/office-insider).</span></span>
- <span data-ttu-id="1025b-108">一般法人向け Office 365 をご利用の場合は、「[一般法人向け Office 365 の先行リリース ビルドをインストールする](https://support.office.com/article/Install-the-First-Release-build-for-Office-365-for-business-customers-4dd8ba40-73c0-4468-b778-c7b744d03ead)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1025b-108">If you're an Office 365 for business customer, see [Install the First Release build for Office 365 for business customers](https://support.office.com/article/Install-the-First-Release-build-for-Office-365-for-business-customers-4dd8ba40-73c0-4468-b778-c7b744d03ead).</span></span>
- <span data-ttu-id="1025b-109">Mac で Office 2016 を実行している場合:</span><span class="sxs-lookup"><span data-stu-id="1025b-109">If you're running Office 2016 on a Mac:</span></span>
    - <span data-ttu-id="1025b-110">Office 2016 for Mac プログラムを起動します。</span><span class="sxs-lookup"><span data-stu-id="1025b-110">Start an Office 2016 for Mac program.</span></span>
    - <span data-ttu-id="1025b-111">[ヘルプ] メニューで [**更新プログラムのチェック**] を選びます。</span><span class="sxs-lookup"><span data-stu-id="1025b-111">Select **Check for Updates** on the Help menu.</span></span>
    - <span data-ttu-id="1025b-112">[Microsoft AutoUpdate] ボックスで、チェック ボックスをオンにして Office Insider プログラムに参加します。</span><span class="sxs-lookup"><span data-stu-id="1025b-112">In the Microsoft AutoUpdate box, check the box to join the Office Insider program.</span></span> 

## <a name="get-the-latest-build"></a><span data-ttu-id="1025b-113">最新ビルドを取得する:</span><span class="sxs-lookup"><span data-stu-id="1025b-113">Get the latest build</span></span>

<span data-ttu-id="1025b-114">Office 2016 の最新ビルドを取得するには:</span><span class="sxs-lookup"><span data-stu-id="1025b-114">To get the latest build of Office 2016:</span></span> 

1. <span data-ttu-id="1025b-115">[ Office 展開ツールのダウンロード](https://www.microsoft.com/download/details.aspx?id=49117) 。</span><span class="sxs-lookup"><span data-stu-id="1025b-115">Download the Office Deployment Tool</span></span> 
2. <span data-ttu-id="1025b-p101">ツールを実行します。これにより、Setup.exe および configuration.xml という 2 つのファイルが抽出されます。</span><span class="sxs-lookup"><span data-stu-id="1025b-p101">Run the tool. This extracts the following two files: Setup.exe and configuration.xml.</span></span>
3. <span data-ttu-id="1025b-118">configuration.xml を[先行リリース構成ファイル](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-Commands-Samples/master/Tools/FirstReleaseConfig/configuration.xml)に置き換えます。</span><span class="sxs-lookup"><span data-stu-id="1025b-118">Replace the configuration.xml file with the [First Release Configuration File](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-Commands-Samples/master/Tools/FirstReleaseConfig/configuration.xml).</span></span>
4. <span data-ttu-id="1025b-119">次のコマンドを管理者として実行します:  `setup.exe /configure configuration.xml`</span><span class="sxs-lookup"><span data-stu-id="1025b-119">Run the following command as an administrator:  `setup.exe /configure configuration.xml`</span></span> 

    > [!NOTE]
    > <span data-ttu-id="1025b-120">このコマンドの実行には時間がかかることがあります (進行状況は表示されません)。</span><span class="sxs-lookup"><span data-stu-id="1025b-120">The command might take a long time to run without indicating progress.</span></span>

<span data-ttu-id="1025b-121">インストール プロセスが完了すると、インストールされている最新の Office アプリケーションがあります。</span><span class="sxs-lookup"><span data-stu-id="1025b-121">When the installation process finishes, you will have the latest Office applications installed.</span></span> <span data-ttu-id="1025b-122">最新のビルドがあることを確認するには、 **ファイル**に移動 > 任意の Office アプリケーションからの**アカウント** です。</span><span class="sxs-lookup"><span data-stu-id="1025b-122">To verify that you have the latest build, go to **File** > **Account** from any Office application.</span></span> <span data-ttu-id="1025b-123">[Office 更新プログラム] に、[(Office Insiders)] ラベルが表示され、その下にバージョン番号が表示されます。</span><span class="sxs-lookup"><span data-stu-id="1025b-123">Under Office Updates, you'll see the (Office Insiders) label above the version number.</span></span>

![Office Insiders のラベルと製品情報を示すスクリーンショット](../images/office-insiders.png)

## <a name="minimum-office-builds-for-office-javascript-api-requirement-sets"></a><span data-ttu-id="1025b-125">Office JavaScript API の要件セットの最小 Office ビルド</span><span class="sxs-lookup"><span data-stu-id="1025b-125">Minimum Office builds for Office JavaScript API requirement sets</span></span>

<span data-ttu-id="1025b-126">API の要件セットの各プラットフォームの最小製品ビルドについては、次をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="1025b-126">For information about the minimum product builds for each platform for the API requirement sets, see the following:</span></span>

- [<span data-ttu-id="1025b-127">Word JavaScript API の要件セット</span><span class="sxs-lookup"><span data-stu-id="1025b-127">Word JavaScript API requirement sets</span></span>](https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets?view=office-js)
- [<span data-ttu-id="1025b-128">Excel JavaScript API の要件セット</span><span class="sxs-lookup"><span data-stu-id="1025b-128">Excel JavaScript API requirement sets</span></span>](https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets?view=office-js)
- [<span data-ttu-id="1025b-129">OneNote JavaScript API の要件セット</span><span class="sxs-lookup"><span data-stu-id="1025b-129">OneNote JavaScript API requirement sets</span></span>](https://docs.microsoft.com/javascript/office/requirement-sets/onenote-api-requirement-sets?view=office-js)
- [<span data-ttu-id="1025b-130">ダイアログ API の要件セット</span><span class="sxs-lookup"><span data-stu-id="1025b-130">Dialog API requirement sets</span></span>](https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets?view=office-js)
- [<span data-ttu-id="1025b-131">Office 共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="1025b-131">Office common API requirement sets</span></span>](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets?view=office-js)
