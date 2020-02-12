---
title: Office の最新バージョンをインストールする
description: Office の最新ビルドを取得するためにオプトインする方法に関する情報。
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: 1f08595ec5d4b7821bf0f2954b306108b0c449bb
ms.sourcegitcommit: d15bca2c12732f8599be2ec4b2adc7c254552f52
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/12/2020
ms.locfileid: "41950671"
---
# <a name="install-the-latest-version-of-office"></a><span data-ttu-id="f4aa6-103">Office の最新バージョンをインストールする</span><span class="sxs-lookup"><span data-stu-id="f4aa6-103">Install the latest version of Office</span></span>

<span data-ttu-id="f4aa6-104">新しい開発者用機能 (現時点ではプレビュー版のものを含む) は、Office の最新ビルドの取得をオプトインしたサブスクライバーに最初に配信されます。</span><span class="sxs-lookup"><span data-stu-id="f4aa6-104">New developer features, including those still in preview, are delivered first to subscribers who opt in to get the latest builds of Office.</span></span>

## <a name="opt-in-to-getting-the-latest-builds"></a><span data-ttu-id="f4aa6-105">最新のビルドを取得するためにオプトインする</span><span class="sxs-lookup"><span data-stu-id="f4aa6-105">Opt in to getting the latest builds</span></span>

<span data-ttu-id="f4aa6-106">Office の最新ビルドを取得するためにオプトインするには、次の操作を行います。</span><span class="sxs-lookup"><span data-stu-id="f4aa6-106">To opt in to getting the latest builds of Office:</span></span>

- <span data-ttu-id="f4aa6-107">Office 365 Solo のサブスクライバーは、「[Office Insider になる](https://products.office.com/office-insider)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f4aa6-107">If you're an Office 365 Home, Personal, or University subscriber, see [Be an Office Insider](https://products.office.com/office-insider).</span></span>
- <span data-ttu-id="f4aa6-108">一般法人向け Office 365 をご利用の場合は、「[一般法人向け Office 365 の先行リリース ビルドをインストールする](https://support.office.com/article/Install-the-First-Release-build-for-Office-365-for-business-customers-4dd8ba40-73c0-4468-b778-c7b744d03ead)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f4aa6-108">If you're an Office 365 for business customer, see [Install the First Release build for Office 365 for business customers](https://support.office.com/article/Install-the-First-Release-build-for-Office-365-for-business-customers-4dd8ba40-73c0-4468-b778-c7b744d03ead).</span></span>
- <span data-ttu-id="f4aa6-109">Mac で Office を実行している場合は、次の操作を行います。</span><span class="sxs-lookup"><span data-stu-id="f4aa6-109">If you're running Office on a Mac:</span></span>
  - <span data-ttu-id="f4aa6-110">Office アプリケーションを起動します。</span><span class="sxs-lookup"><span data-stu-id="f4aa6-110">Start an Office application.</span></span>
  - <span data-ttu-id="f4aa6-111">[ヘルプ] メニューで [**更新プログラムのチェック**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="f4aa6-111">Select **Check for Updates** on the Help menu.</span></span>
  - <span data-ttu-id="f4aa6-112">[Microsoft AutoUpdate] ボックスで、チェック ボックスをオンにして Office Insider プログラムに参加します。</span><span class="sxs-lookup"><span data-stu-id="f4aa6-112">In the Microsoft AutoUpdate box, check the box to join the Office Insider program.</span></span>

## <a name="get-the-latest-build"></a><span data-ttu-id="f4aa6-113">最新ビルドを取得する</span><span class="sxs-lookup"><span data-stu-id="f4aa6-113">Get the latest build</span></span>

<span data-ttu-id="f4aa6-114">Office の最新ビルドを取得するには、次の操作を行います。</span><span class="sxs-lookup"><span data-stu-id="f4aa6-114">To get the latest build of Office:</span></span>

1. <span data-ttu-id="f4aa6-115">[Office 展開ツール](https://www.microsoft.com/download/details.aspx?id=49117)をダウンロードします。</span><span class="sxs-lookup"><span data-stu-id="f4aa6-115">Download the [Office Deployment Tool](https://www.microsoft.com/download/details.aspx?id=49117).</span></span>
2. <span data-ttu-id="f4aa6-p101">ツールを実行します。これにより、Setup.exe および configuration.xml という 2 つのファイルが抽出されます。</span><span class="sxs-lookup"><span data-stu-id="f4aa6-p101">Run the tool. This extracts the following two files: Setup.exe and configuration.xml.</span></span>
3. <span data-ttu-id="f4aa6-118">configuration.xml を[先行リリース構成ファイル](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-Commands-Samples/master/Tools/FirstReleaseConfig/configuration.xml)に置き換えます。</span><span class="sxs-lookup"><span data-stu-id="f4aa6-118">Replace the configuration.xml file with the [First Release Configuration File](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-Commands-Samples/master/Tools/FirstReleaseConfig/configuration.xml).</span></span>
4. <span data-ttu-id="f4aa6-119">次のコマンドを管理者として実行します: `setup.exe /configure configuration.xml`</span><span class="sxs-lookup"><span data-stu-id="f4aa6-119">Run the following command as an administrator:  `setup.exe /configure configuration.xml`</span></span>

> [!NOTE]
> <span data-ttu-id="f4aa6-120">このコマンドの実行には時間がかかることがあります (進行状況は表示されません)。</span><span class="sxs-lookup"><span data-stu-id="f4aa6-120">The command might take a long time to run without indicating progress.</span></span>

<span data-ttu-id="f4aa6-121">インストール処理の完了時点で、最新の Office アプリケーションがインストールされています。</span><span class="sxs-lookup"><span data-stu-id="f4aa6-121">When the installation process finishes, you will have the latest Office applications installed.</span></span> <span data-ttu-id="f4aa6-122">最新のビルドであることを確認するには、任意の Office アプリケーションから **[ファイル]**、**[アカウント]** の順に移動します。</span><span class="sxs-lookup"><span data-stu-id="f4aa6-122">To verify that you have the latest build, go to **File** > **Account** from any Office application.</span></span> <span data-ttu-id="f4aa6-123">[Office 更新プログラム] に、[(Office Insiders)] ラベルが表示され、その下にバージョン番号が表示されます。</span><span class="sxs-lookup"><span data-stu-id="f4aa6-123">Under Office Updates, you'll see the (Office Insiders) label above the version number.</span></span>

![Office Insiders のラベルと製品情報を示すスクリーンショット](../images/office-insiders-label.png)

## <a name="minimum-office-builds-for-office-javascript-api-requirement-sets"></a><span data-ttu-id="f4aa6-125">Office JavaScript API の要件セットの最小 Office ビルド</span><span class="sxs-lookup"><span data-stu-id="f4aa6-125">Minimum Office builds for Office JavaScript API requirement sets</span></span>

<span data-ttu-id="f4aa6-126">API の要件セットの各プラットフォームの最小製品ビルドについては、次をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="f4aa6-126">For information about the minimum product builds for each platform for the API requirement sets, see the following:</span></span>

- [<span data-ttu-id="f4aa6-127">Excel JavaScript API の要件セット</span><span class="sxs-lookup"><span data-stu-id="f4aa6-127">Excel JavaScript API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets)
- [<span data-ttu-id="f4aa6-128">OneNote JavaScript API の要件セット</span><span class="sxs-lookup"><span data-stu-id="f4aa6-128">OneNote JavaScript API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets)
- [<span data-ttu-id="f4aa6-129">Outlook JavaScript API の要件セット</span><span class="sxs-lookup"><span data-stu-id="f4aa6-129">Outlook JavaScript API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)
- [<span data-ttu-id="f4aa6-130">PowerPoint JavaScript API の要件セット</span><span class="sxs-lookup"><span data-stu-id="f4aa6-130">PowerPoint JavaScript API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets)
- [<span data-ttu-id="f4aa6-131">Word JavaScript API の要件セット</span><span class="sxs-lookup"><span data-stu-id="f4aa6-131">Word JavaScript API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets)
- [<span data-ttu-id="f4aa6-132">ダイアログ API の要件セット</span><span class="sxs-lookup"><span data-stu-id="f4aa6-132">Dialog API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets)
- [<span data-ttu-id="f4aa6-133">Office 共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="f4aa6-133">Office Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
