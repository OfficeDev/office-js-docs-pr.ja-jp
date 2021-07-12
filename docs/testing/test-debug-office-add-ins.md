---
title: Office アドインのテストとデバッグ
description: Office アドインのテストとデバッグを行う方法について説明します。
ms.date: 05/19/2021
localization_priority: Priority
ms.openlocfilehash: 5df42a6c22325528eaaf2dcde28fddbfd3a211fb
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349358"
---
# <a name="test-and-debug-office-add-ins"></a><span data-ttu-id="c2d8b-103">Office アドインのテストとデバッグ</span><span class="sxs-lookup"><span data-stu-id="c2d8b-103">Test and debug Office Add-ins</span></span>

<span data-ttu-id="c2d8b-104">この記事では、Office アドインのテスト、デバッグ、トラブルシューティングに関するガイダンスを示します。</span><span class="sxs-lookup"><span data-stu-id="c2d8b-104">This article contains guidance about testing, debugging, and troubleshooting issues with Office Add-ins.</span></span>

## <a name="test-cross-platform-and-for-multiple-versions-of-office"></a><span data-ttu-id="c2d8b-105">クロスプラットフォームおよび複数バージョンの Office をテストする</span><span class="sxs-lookup"><span data-stu-id="c2d8b-105">Test cross-platform and for multiple versions of Office</span></span>

<span data-ttu-id="c2d8b-106">Office アドインは主要なプラットフォームで実行されるため、ユーザーが Office を実行している可能性のあるすべてのプラットフォームでアドインをテストする必要があります。</span><span class="sxs-lookup"><span data-stu-id="c2d8b-106">Office Add-ins run across major platforms, so you need to test an add-in in all the platforms where your users might be running Office.</span></span> <span data-ttu-id="c2d8b-107">これには通常、Office on the web、Office on Windows (サブスクリプションと 1 回限りの購入の両方)、Office on Mac、Office on iOS、および (Outlook アドインの場合) Office on Android が含まれます。</span><span class="sxs-lookup"><span data-stu-id="c2d8b-107">This usually includes Office on the web, Office on Windows (both subscription and one-time purchase), Office on Mac, Office on iOS, and (for Outlook add-ins) Office on Android.</span></span> <span data-ttu-id="c2d8b-108">ただし、一部のプラットフォームで作業しているユーザーがいないことを確認できる場合もあります。</span><span class="sxs-lookup"><span data-stu-id="c2d8b-108">However, there may be some situations in which you can be sure that none of your users will be working on some platforms.</span></span> <span data-ttu-id="c2d8b-109">たとえば、ユーザーが Windows コンピューターとサブスクリプション Office を使用する必要がある会社のアドインを作成する場合、Office on Mac や 1 回限りの購入の Windows をテストする必要はありません。</span><span class="sxs-lookup"><span data-stu-id="c2d8b-109">For example, if you are making an add-in for a company that requires its users to work with Windows computers and subscription Office, then you don't need to test for Office on Mac or one-time purchase Windows.</span></span>

> [!NOTE]
> <span data-ttu-id="c2d8b-110">Windows コンピューターでは、Windows と Office のバージョンによって、アドインが使用するブラウザー コントロールが決まります。詳細については、「[Office アドインによって使用されるブラウザー](../concepts/browsers-used-by-office-web-add-ins.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c2d8b-110">On Windows computers, the version of Windows and Office will determine which browser control is used by add-ins. For more information, see [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="c2d8b-111">AppSource を通じて販売されるアドインは、すべてのプラットフォームでのテストを含む検証プロセスを経ます。</span><span class="sxs-lookup"><span data-stu-id="c2d8b-111">Add-ins marketed through AppSource go through a validation process that includes testing on all platforms.</span></span> <span data-ttu-id="c2d8b-112">さらに、アドインは、Microsoft Edge (Chromium ベースの WebView2)、Chrome、Safari など、すべての主要な最新のブラウザーで Office on the web 用にテストされています。</span><span class="sxs-lookup"><span data-stu-id="c2d8b-112">In addition, add-ins are tested for Office on the web with all major modern browsers, including Microsoft Edge (Chromium-based WebView2), Chrome, and Safari.</span></span> <span data-ttu-id="c2d8b-113">したがって、AppSource に送信する前に、これらのプラットフォームとブラウザーでテストする必要があります。</span><span class="sxs-lookup"><span data-stu-id="c2d8b-113">Accordingly, you should test on these platforms and browsers before you submit to AppSource.</span></span> <span data-ttu-id="c2d8b-114">検証の詳細については、「[コマーシャル マーケットプレースの認定ポリシー](/legal/marketplace/certification-policies)」、特に[セクション 1120.3](/legal/marketplace/certification-policies#11203-functionality)、および [Office アドイン アプリケーションと可用性のページ](../overview/office-add-in-availability.md)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c2d8b-114">For more information about validation, see [Commercial marketplace certification policies](/legal/marketplace/certification-policies), especially [section 1120.3](/legal/marketplace/certification-policies#11203-functionality), and the [Office Add-in application and availability page](../overview/office-add-in-availability.md).</span></span>
>
> <span data-ttu-id="c2d8b-115">AppSource は、Office on the web でアドインをテストするために、Internet Explorer または Microsoft Edge の以前のバージョン (WebView1) を使用しません。</span><span class="sxs-lookup"><span data-stu-id="c2d8b-115">AppSource does not use Internet Explorer or the legacy version of Microsoft Edge (WebView1) to test add-ins in Office on the web.</span></span> <span data-ttu-id="c2d8b-116">ただし、かなりの数のユーザーがこれら 2 つのブラウザーを使用して Office on the web を開く場合は、それらのブラウザーでテストする必要があります。</span><span class="sxs-lookup"><span data-stu-id="c2d8b-116">But if a significant number of your users will use these two browsers to open Office on the web, then you should test with them.</span></span> <span data-ttu-id="c2d8b-117">詳細については、「[Internet Explorer 11 のサポート](../develop/support-ie-11.md)」および「[Microsoft Edge の問題のトラブルシューティング](../concepts/browsers-used-by-office-web-add-ins.md#troubleshooting-microsoft-edge-issues)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c2d8b-117">For more information, see [Support Internet Explorer 11](../develop/support-ie-11.md) and [Troubleshooting Microsoft Edge issues](../concepts/browsers-used-by-office-web-add-ins.md#troubleshooting-microsoft-edge-issues).</span></span> <span data-ttu-id="c2d8b-118">Office は引き続きアドイン用にこれらのブラウザーをサポートしているため、アドインの実行時にバグが発生したと思われる場合は、[office-js](https://github.com/OfficeDev/office-js/issues/new/choose) リポジトリの問題を作成してください。</span><span class="sxs-lookup"><span data-stu-id="c2d8b-118">Office still supports these browsers for add-ins, so if you think you've encountered a bug in how add-ins run in them, please create an issue for the [office-js](https://github.com/OfficeDev/office-js/issues/new/choose) repo.</span></span>

## <a name="sideload-an-office-add-in-for-testing"></a><span data-ttu-id="c2d8b-119">テスト用に Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="c2d8b-119">Sideload an Office Add-in for testing</span></span>

<span data-ttu-id="c2d8b-p104">サイドロードを使用すると、最初にアドイン カタログに置かなくても、テスト用に Office アドインをインストールすることができます。アドインをサイドロードする手順は、プラットフォームによって異なり、場合によっては、製品によっても異なります。次のそれぞれの記事では、特定のプラットフォームまたは特定の製品の Office アドインをサイドロードする方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="c2d8b-p104">You can use sideloading to install an Office Add-in for testing without having to first put it in an add-in catalog. The procedure for sideloading an add-in varies by platform, and in some cases, by product as well. The following articles each describe how to sideload Office Add-ins on a specific platform or within a specific product.</span></span>

- [<span data-ttu-id="c2d8b-123">Windows で Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="c2d8b-123">Sideload Office Add-ins on Windows</span></span>](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)

- [<span data-ttu-id="c2d8b-124">Office on the web で Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="c2d8b-124">Sideload Office Add-ins in Office on the web</span></span>](sideload-office-add-ins-for-testing.md)

- [<span data-ttu-id="c2d8b-125">iPad と Mac で Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="c2d8b-125">Sideload Office Add-ins on iPad and Mac</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)

- [<span data-ttu-id="c2d8b-126">テスト用に Outlook アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="c2d8b-126">Sideload Outlook add-ins for testing</span></span>](../outlook/sideload-outlook-add-ins-for-testing.md)

## <a name="debug-an-office-add-in"></a><span data-ttu-id="c2d8b-127">Office アドインのデバッグ</span><span class="sxs-lookup"><span data-stu-id="c2d8b-127">Debug an Office Add-in</span></span>

<span data-ttu-id="c2d8b-p105">Office アドインをデバッグする手順も、プラットフォームによって異なります。次のそれぞれの記事では、特定のプラットフォームで Office アドインをデバッグする方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="c2d8b-p105">The procedure for debugging an Office Add-in varies by platform as well. Each of the following articles describes how to debug Office Add-ins on a specific platform.</span></span>

- [<span data-ttu-id="c2d8b-130">(Windows で) 作業ウィンドウからデバッガーをアタッチする</span><span class="sxs-lookup"><span data-stu-id="c2d8b-130">Attach a debugger from the task pane (on Windows)</span></span>](attach-debugger-from-task-pane.md)

- [<span data-ttu-id="c2d8b-131">Windows 10 で F12 開発者ツールを使用してアドインをデバッグする</span><span class="sxs-lookup"><span data-stu-id="c2d8b-131">Debug add-ins using F12 developer tools on Windows 10</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)

- [<span data-ttu-id="c2d8b-132">Office on the web でアドインをデバッグする</span><span class="sxs-lookup"><span data-stu-id="c2d8b-132">Debug add-ins in Office on the web</span></span>](debug-add-ins-in-office-online.md)

- [<span data-ttu-id="c2d8b-133">iPad と Mac で Office アドインをデバッグする</span><span class="sxs-lookup"><span data-stu-id="c2d8b-133">Debug Office Add-ins on iPad and Mac</span></span>](debug-office-add-ins-on-ipad-and-mac.md)

- [<span data-ttu-id="c2d8b-134">Visual Studio Code 用 Microsoft Office アドイン デバッガー拡張機能</span><span class="sxs-lookup"><span data-stu-id="c2d8b-134">Microsoft Office Add-in Debugger Extension for Visual Studio Code</span></span>](debug-with-vs-extension.md)

## <a name="validate-an-office-add-in-manifest"></a><span data-ttu-id="c2d8b-135">Office アドイン マニフェストの検証</span><span class="sxs-lookup"><span data-stu-id="c2d8b-135">Validate an Office Add-in manifest</span></span>

<span data-ttu-id="c2d8b-136">Office アドインを記述するマニフェスト ファイルを検証し、マニフェスト ファイルの問題のトラブルシューティングを行う方法については、「[マニフェストの問題を検証し、トラブルシューティングを行う](troubleshoot-manifest.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c2d8b-136">For information about how to validate the manifest file that describes your Office Add-in and troubleshoot issues with the manifest file, see [Validate and troubleshoot issues with your manifest](troubleshoot-manifest.md).</span></span>

## <a name="troubleshoot-user-errors"></a><span data-ttu-id="c2d8b-137">ユーザーのエラーのトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="c2d8b-137">Troubleshoot user errors</span></span>

<span data-ttu-id="c2d8b-138">よくある Office アドインの問題の解決方法については、「[Office アドインでのユーザー エラーのトラブルシューティング](testing-and-troubleshooting.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c2d8b-138">For information about how to resolve common issues that users may encounter with your Office Add-in, see [Troubleshoot user errors with Office Add-ins](testing-and-troubleshooting.md).</span></span>
