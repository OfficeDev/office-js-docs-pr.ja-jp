---
title: Office 2016 の最新バージョンをインストールする
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 98dc69a7971a94b96bc3f7304fc7905f31013a87
ms.sourcegitcommit: 4de2a1b62ccaa8e51982e95537fc9f52c0c5e687
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/10/2018
ms.locfileid: "22925235"
---
# <a name="install-the-latest-version-of-office-2016"></a><span data-ttu-id="48334-102">Office 2016 の最新バージョンをインストールする</span><span class="sxs-lookup"><span data-stu-id="48334-102">Install the latest version of Office 2016</span></span>

<span data-ttu-id="48334-103">新しい開発者用機能 (現時点ではプレビュー版のものを含む) は、Office の最新ビルドの取得をオプトインしたサブスクライバーに最初に配信されます。</span><span class="sxs-lookup"><span data-stu-id="48334-103">New developer features, including those still in preview, are delivered first to subscribers who opt in to get the latest builds of Office.</span></span> 

## <a name="opt-in-to-getting-the-latest-builds"></a><span data-ttu-id="48334-104">最新のビルドを取得するためにオプトインする</span><span class="sxs-lookup"><span data-stu-id="48334-104">Opt in to getting the latest builds</span></span>

<span data-ttu-id="48334-105">Office 2016 の最新ビルドを取得するためにオプトインするには:</span><span class="sxs-lookup"><span data-stu-id="48334-105">To opt in to getting the latest builds of Office 2016:</span></span> 

- <span data-ttu-id="48334-106">Office 365 Home、Personal、または University のサブスクライバーは、「[Office Insider プログラム](https://products.office.com/office-insider)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="48334-106">If you're an Office 365 Home, Personal, or University subscriber, see [Be an Office Insider](https://products.office.com/office-insider).</span></span>
- <span data-ttu-id="48334-107">一般法人向け Office 365 をご利用の場合は、「[一般法人向け Office 365 の先行リリース ビルドをインストールする](https://support.office.com/article/Install-the-First-Release-build-for-Office-365-for-business-customers-4dd8ba40-73c0-4468-b778-c7b744d03ead)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="48334-107">If you're an Office 365 for business customer, see [Install the First Release build for Office 365 for business customers](https://support.office.com/article/Install-the-First-Release-build-for-Office-365-for-business-customers-4dd8ba40-73c0-4468-b778-c7b744d03ead).</span></span>
- <span data-ttu-id="48334-108">Mac で Office 2016 を実行している場合:</span><span class="sxs-lookup"><span data-stu-id="48334-108">If you're running Office 2016 on a Mac:</span></span>
    - <span data-ttu-id="48334-109">Office 2016 for Mac プログラムを起動します。</span><span class="sxs-lookup"><span data-stu-id="48334-109">Start an Office 2016 for Mac program.</span></span>
    - <span data-ttu-id="48334-110">[ヘルプ] メニューで [**更新プログラムのチェック**] を選びます。</span><span class="sxs-lookup"><span data-stu-id="48334-110">Select **Check for Updates** on the Help menu.</span></span>
    - <span data-ttu-id="48334-111">[Microsoft AutoUpdate] ボックスで、チェック ボックスをオンにして Office Insider プログラムに参加します。</span><span class="sxs-lookup"><span data-stu-id="48334-111">In the Microsoft AutoUpdate box, check the box to join the Office Insider program.</span></span> 

## <a name="get-the-latest-build"></a><span data-ttu-id="48334-112">最新ビルドを取得する:</span><span class="sxs-lookup"><span data-stu-id="48334-112">Get the latest build</span></span>

<span data-ttu-id="48334-113">Office 2016 の最新ビルドを取得するには:</span><span class="sxs-lookup"><span data-stu-id="48334-113">To get the latest build of Office 2016:</span></span> 

1. <span data-ttu-id="48334-114">[Office 2016 展開ツール](https://www.microsoft.com/download/details.aspx?id=49117)をダウンロードします。</span><span class="sxs-lookup"><span data-stu-id="48334-114">Download the [Office 2016 Deployment Tool](https://www.microsoft.com/download/details.aspx?id=49117).</span></span> 
2. <span data-ttu-id="48334-p101">ツールを実行します。これにより、Setup.exe および configuration.xml という 2 つのファイルが抽出されます。</span><span class="sxs-lookup"><span data-stu-id="48334-p101">Run the tool. This extracts the following two files: Setup.exe and configuration.xml.</span></span>
3. <span data-ttu-id="48334-117">configuration.xml を[先行リリース構成ファイル](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-Commands-Samples/master/Tools/FirstReleaseConfig/configuration.xml)に置き換えます。</span><span class="sxs-lookup"><span data-stu-id="48334-117">Replace the configuration.xml file with the [First Release Configuration File](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-Commands-Samples/master/Tools/FirstReleaseConfig/configuration.xml).</span></span>
4. <span data-ttu-id="48334-118">次のコマンドを管理者として実行します: `setup.exe /configure configuration.xml`</span><span class="sxs-lookup"><span data-stu-id="48334-118">Run the following command as an administrator:  `setup.exe /configure configuration.xml`</span></span> 

    > [!NOTE]
    > <span data-ttu-id="48334-119">このコマンドの実行には時間がかかることがあります (進行状況は表示されません)。</span><span class="sxs-lookup"><span data-stu-id="48334-119">The command might take a long time to run without indicating progress.</span></span>

<span data-ttu-id="48334-p102">インストール処理の完了時点で、最新の Office 2016 アプリケーションがインストールされています。最新のビルドであることを確認するには、任意の Office アプリケーションから **[ファイル]**  >  **[アカウント]** に移動します。[Office 更新プログラム] に、[(Office Insiders)] ラベルが表示され、その下にバージョン番号が表示されます。</span><span class="sxs-lookup"><span data-stu-id="48334-p102">When the installation process finishes, you will have the latest Office 2016 applications installed. To verify that you have the latest build, go to **File** > **Account** from any Office application. Under Office Updates, you'll see the (Office Insiders) label above the version number.</span></span>

![Office Insiders のラベルと製品情報を示すスクリーンショット](../images/office-insiders.png)

## <a name="minimum-office-builds-for-office-javascript-api-requirement-sets"></a><span data-ttu-id="48334-124">Office JavaScript API の要件セットの最小 Office ビルド</span><span class="sxs-lookup"><span data-stu-id="48334-124">Minimum Office builds for Office JavaScript API requirement sets</span></span>

<span data-ttu-id="48334-125">API の要件セットの各プラットフォームの最小製品ビルドについては、次をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="48334-125">For information about the minimum product builds for each platform for the API requirement sets, see the following:</span></span>

- [<span data-ttu-id="48334-126">Word JavaScript API の要件セット</span><span class="sxs-lookup"><span data-stu-id="48334-126">Word JavaScript API requirement sets</span></span>](https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets)
- [<span data-ttu-id="48334-127">Excel JavaScript API の要件セット</span><span class="sxs-lookup"><span data-stu-id="48334-127">Excel JavaScript API requirement sets</span></span>](https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets)
- [<span data-ttu-id="48334-128">OneNote JavaScript API の要件セット</span><span class="sxs-lookup"><span data-stu-id="48334-128">OneNote JavaScript API requirement sets</span></span>](https://dev.office.com/reference/add-ins/requirement-sets/onenote-api-requirement-sets)
- [<span data-ttu-id="48334-129">ダイアログ API の要件セット</span><span class="sxs-lookup"><span data-stu-id="48334-129">Dialog API requirement sets</span></span>](https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets)
- [<span data-ttu-id="48334-130">Office 共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="48334-130">Office common API requirement sets</span></span>](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets)
