---
title: Office の最新バージョンをインストールする
description: Office の最新ビルドを取得するためにオプトインする方法に関する情報。
ms.date: 04/16/2020
localization_priority: Normal
ms.openlocfilehash: b9428cc67160e0680bab5a36438bc1a0dbb3ac17
ms.sourcegitcommit: 803587b324fc8038721709d7db5664025cf03c6b
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/17/2020
ms.locfileid: "43547065"
---
# <a name="install-the-latest-version-of-office"></a><span data-ttu-id="613bd-103">Office の最新バージョンをインストールする</span><span class="sxs-lookup"><span data-stu-id="613bd-103">Install the latest version of Office</span></span>

<span data-ttu-id="613bd-104">新しい開発者用機能 (現時点ではプレビュー版のものを含む) は、Office の最新ビルドの取得をオプトインしたサブスクライバーに最初に配信されます。</span><span class="sxs-lookup"><span data-stu-id="613bd-104">New developer features, including those still in preview, are delivered first to subscribers who opt in to get the latest builds of Office.</span></span>

## <a name="opt-in-to-getting-the-latest-builds"></a><span data-ttu-id="613bd-105">最新のビルドを取得するためにオプトインする</span><span class="sxs-lookup"><span data-stu-id="613bd-105">Opt in to getting the latest builds</span></span>

<span data-ttu-id="613bd-106">Office の最新ビルドを取得するためにオプトインするには、次の操作を行います。</span><span class="sxs-lookup"><span data-stu-id="613bd-106">To opt in to getting the latest builds of Office:</span></span>

- <span data-ttu-id="613bd-107">Office 365 Solo のサブスクライバーは、「[Office Insider になる](https://insider.office.com)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="613bd-107">If you're an Office 365 Home, Personal, or University subscriber, see [Be an Office Insider](https://insider.office.com).</span></span>
- <span data-ttu-id="613bd-108">一般法人向け Office 365 をご利用の場合は、「[一般法人向け Office 365 の先行リリース ビルドをインストールする](https://support.office.com/article/Install-the-First-Release-build-for-Office-365-for-business-customers-4dd8ba40-73c0-4468-b778-c7b744d03ead)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="613bd-108">If you're an Office 365 for business customer, see [Install the First Release build for Office 365 for business customers](https://support.office.com/article/Install-the-First-Release-build-for-Office-365-for-business-customers-4dd8ba40-73c0-4468-b778-c7b744d03ead).</span></span>
- <span data-ttu-id="613bd-109">Mac で Office を実行している場合は、次の操作を行います。</span><span class="sxs-lookup"><span data-stu-id="613bd-109">If you're running Office on a Mac:</span></span>
  - <span data-ttu-id="613bd-110">Office アプリケーションを起動します。</span><span class="sxs-lookup"><span data-stu-id="613bd-110">Start an Office application.</span></span>
  - <span data-ttu-id="613bd-111">[ヘルプ] メニューで [**更新プログラムのチェック**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="613bd-111">Select **Check for Updates** on the Help menu.</span></span>
  - <span data-ttu-id="613bd-112">[Microsoft AutoUpdate] ボックスで、チェック ボックスをオンにして Office Insider プログラムに参加します。</span><span class="sxs-lookup"><span data-stu-id="613bd-112">In the Microsoft AutoUpdate box, check the box to join the Office Insider program.</span></span>

## <a name="get-the-latest-build"></a><span data-ttu-id="613bd-113">最新ビルドを取得する</span><span class="sxs-lookup"><span data-stu-id="613bd-113">Get the latest build</span></span>

<span data-ttu-id="613bd-114">Office の最新ビルドを取得するには、次の操作を行います。</span><span class="sxs-lookup"><span data-stu-id="613bd-114">To get the latest build of Office:</span></span>

1. <span data-ttu-id="613bd-115">[Office 展開ツール](https://www.microsoft.com/download/details.aspx?id=49117)をダウンロードします。</span><span class="sxs-lookup"><span data-stu-id="613bd-115">Download the [Office Deployment Tool](https://www.microsoft.com/download/details.aspx?id=49117).</span></span>
2. <span data-ttu-id="613bd-p101">ツールを実行します。これにより、Setup.exe および configuration.xml という 2 つのファイルが抽出されます。</span><span class="sxs-lookup"><span data-stu-id="613bd-p101">Run the tool. This extracts the following two files: Setup.exe and configuration.xml.</span></span>
3. <span data-ttu-id="613bd-118">configuration.xml を[先行リリース構成ファイル](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-Commands-Samples/master/Tools/FirstReleaseConfig/configuration.xml)に置き換えます。</span><span class="sxs-lookup"><span data-stu-id="613bd-118">Replace the configuration.xml file with the [First Release Configuration File](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-Commands-Samples/master/Tools/FirstReleaseConfig/configuration.xml).</span></span>
4. <span data-ttu-id="613bd-119">次のコマンドを管理者として実行します: `setup.exe /configure configuration.xml`</span><span class="sxs-lookup"><span data-stu-id="613bd-119">Run the following command as an administrator:  `setup.exe /configure configuration.xml`</span></span>

> [!NOTE]
> <span data-ttu-id="613bd-120">このコマンドの実行には時間がかかることがあります (進行状況は表示されません)。</span><span class="sxs-lookup"><span data-stu-id="613bd-120">The command might take a long time to run without indicating progress.</span></span>

<span data-ttu-id="613bd-121">インストール処理の完了時点で、最新の Office アプリケーションがインストールされています。</span><span class="sxs-lookup"><span data-stu-id="613bd-121">When the installation process finishes, you will have the latest Office applications installed.</span></span> <span data-ttu-id="613bd-122">最新のビルドであることを確認するには、任意の Office アプリケーションから **[ファイル]**、**[アカウント]** の順に移動します。</span><span class="sxs-lookup"><span data-stu-id="613bd-122">To verify that you have the latest build, go to **File** > **Account** from any Office application.</span></span> <span data-ttu-id="613bd-123">[Office 更新プログラム] に、[(Office Insiders)] ラベルが表示され、その下にバージョン番号が表示されます。</span><span class="sxs-lookup"><span data-stu-id="613bd-123">Under Office Updates, you'll see the (Office Insiders) label above the version number.</span></span>

![Office Insiders のラベルと製品情報を示すスクリーンショット](../images/office-insiders-label.png)

## <a name="minimum-office-builds-for-office-javascript-api-requirement-sets"></a><span data-ttu-id="613bd-125">Office JavaScript API の要件セットの最小 Office ビルド</span><span class="sxs-lookup"><span data-stu-id="613bd-125">Minimum Office builds for Office JavaScript API requirement sets</span></span>

<span data-ttu-id="613bd-126">API の要件セットの各プラットフォームの最小製品ビルドについては、次をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="613bd-126">For information about the minimum product builds for each platform for the API requirement sets, see the following:</span></span>

- [<span data-ttu-id="613bd-127">Excel JavaScript API の要件セット</span><span class="sxs-lookup"><span data-stu-id="613bd-127">Excel JavaScript API requirement sets</span></span>](../reference/requirement-sets/excel-api-requirement-sets.md)
- [<span data-ttu-id="613bd-128">OneNote JavaScript API の要件セット</span><span class="sxs-lookup"><span data-stu-id="613bd-128">OneNote JavaScript API requirement sets</span></span>](../reference/requirement-sets/onenote-api-requirement-sets.md)
- [<span data-ttu-id="613bd-129">Outlook JavaScript API の要件セット</span><span class="sxs-lookup"><span data-stu-id="613bd-129">Outlook JavaScript API requirement sets</span></span>](../reference/requirement-sets/outlook-api-requirement-sets.md)
- [<span data-ttu-id="613bd-130">PowerPoint JavaScript API の要件セット</span><span class="sxs-lookup"><span data-stu-id="613bd-130">PowerPoint JavaScript API requirement sets</span></span>](../reference/requirement-sets/powerpoint-api-requirement-sets.md)
- [<span data-ttu-id="613bd-131">Word JavaScript API の要件セット</span><span class="sxs-lookup"><span data-stu-id="613bd-131">Word JavaScript API requirement sets</span></span>](../reference/requirement-sets/word-api-requirement-sets.md)
- [<span data-ttu-id="613bd-132">ダイアログ API の要件セット</span><span class="sxs-lookup"><span data-stu-id="613bd-132">Dialog API requirement sets</span></span>](../reference/requirement-sets/dialog-api-requirement-sets.md)
- [<span data-ttu-id="613bd-133">Office 共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="613bd-133">Office Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
