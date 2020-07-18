---
title: Office アドインのテストとデバッグ
description: Office アドインのテストとデバッグを行う方法について説明します。
ms.date: 06/17/2020
localization_priority: Priority
ms.openlocfilehash: 526204fe94d4c97ce7e1e0bc9ac2a212f69611d3
ms.sourcegitcommit: 472b81642e9eb5fb2a55cd98a7b0826d37eb7f73
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/17/2020
ms.locfileid: "45159249"
---
# <a name="test-and-debug-office-add-ins"></a><span data-ttu-id="f1e92-103">Office アドインのテストとデバッグ</span><span class="sxs-lookup"><span data-stu-id="f1e92-103">Test and debug Office Add-ins</span></span>

<span data-ttu-id="f1e92-104">このセクションでは、Office アドインのテスト、デバッグ、トラブルシューティングに関するガイダンスを示します。</span><span class="sxs-lookup"><span data-stu-id="f1e92-104">This section contains guidance about testing, debugging, and troubleshooting issues with Office Add-ins.</span></span>

## <a name="sideload-an-office-add-in-for-testing"></a><span data-ttu-id="f1e92-105">テスト用に Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="f1e92-105">Sideload an Office Add-in for testing</span></span>

<span data-ttu-id="f1e92-p101">サイドロードを使用すると、最初にアドイン カタログに置かなくても、テスト用に Office アドインをインストールすることができます。 アドインをサイドロードする手順は、プラットフォームによって異なり、場合によっては、製品によっても異なります。 次のそれぞれの記事では、特定のプラットフォームまたは特定の製品の Office アドインをサイドロードする方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="f1e92-p101">You can use sideloading to install an Office Add-in for testing without having to first put it in an add-in catalog. The procedure for sideloading an add-in varies by platform, and in some cases, by product as well. The following articles each describe how to sideload Office Add-ins on a specific platform or within a specific product:</span></span>

- [<span data-ttu-id="f1e92-109">Windows で Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="f1e92-109">Sideload Office Add-ins on Windows</span></span>](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)

- [<span data-ttu-id="f1e92-110">Office on the web で Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="f1e92-110">Sideload Office Add-ins in Office on the web</span></span>](sideload-office-add-ins-for-testing.md)

- [<span data-ttu-id="f1e92-111">iPad と Mac で Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="f1e92-111">Sideload Office Add-ins on iPad and Mac</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)

- [<span data-ttu-id="f1e92-112">テスト用に Outlook アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="f1e92-112">Sideload Outlook add-ins for testing</span></span>](../outlook/sideload-outlook-add-ins-for-testing.md)

## <a name="debug-an-office-add-in"></a><span data-ttu-id="f1e92-113">Office アドインのデバッグ</span><span class="sxs-lookup"><span data-stu-id="f1e92-113">Debug an Office Add-in</span></span>

<span data-ttu-id="f1e92-p102">Office アドインをデバッグする手順も、プラットフォームによって異なります。 次のそれぞれの記事では、特定のプラットフォームで Office アドインをデバッグする方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="f1e92-p102">The procedure for debugging an Office Add-in varies by platform as well. Each of the following articles describes how to debug Office Add-ins on a specific platform:</span></span>

- [<span data-ttu-id="f1e92-116">(Windows で) 作業ウィンドウからデバッガーをアタッチする</span><span class="sxs-lookup"><span data-stu-id="f1e92-116">Attach a debugger from the task pane (on Windows)</span></span>](attach-debugger-from-task-pane.md)

- [<span data-ttu-id="f1e92-117">Windows 10 で F12 開発者ツールを使用してアドインをデバッグする</span><span class="sxs-lookup"><span data-stu-id="f1e92-117">Debug add-ins using F12 developer tools on Windows 10</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)

- [<span data-ttu-id="f1e92-118">Office on the web でアドインをデバッグする</span><span class="sxs-lookup"><span data-stu-id="f1e92-118">Debug add-ins in Office on the web</span></span>](debug-add-ins-in-office-online.md)

- [<span data-ttu-id="f1e92-119">iPad と Mac で Office アドインをデバッグする</span><span class="sxs-lookup"><span data-stu-id="f1e92-119">Debug Office Add-ins on iPad and Mac</span></span>](debug-office-add-ins-on-ipad-and-mac.md)

- [<span data-ttu-id="f1e92-120">Visual Studio Code 用 Microsoft Office アドイン デバッガー拡張機能</span><span class="sxs-lookup"><span data-stu-id="f1e92-120">Microsoft Office Add-in Debugger Extension for Visual Studio Code</span></span>](debug-with-vs-extension.md)

## <a name="validate-an-office-add-in-manifest"></a><span data-ttu-id="f1e92-121">Office アドイン マニフェストの検証</span><span class="sxs-lookup"><span data-stu-id="f1e92-121">Validate an Office Add-in manifest</span></span>

<span data-ttu-id="f1e92-122">Office アドインを記述するマニフェスト ファイルを検証し、マニフェスト ファイルの問題のトラブルシューティングを行う方法については、「[マニフェストの問題を検証し、トラブルシューティングを行う](troubleshoot-manifest.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f1e92-122">For information about how to validate the manifest file that describes your Office Add-in and troubleshoot issues with the manifest file, see [Validate and troubleshoot issues with your manifest](troubleshoot-manifest.md).</span></span>

## <a name="troubleshoot-user-errors"></a><span data-ttu-id="f1e92-123">ユーザーのエラーのトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="f1e92-123">Troubleshoot user errors</span></span>

<span data-ttu-id="f1e92-124">よくある Office アドインの問題の解決方法については、「[Office アドインでのユーザー エラーのトラブルシューティング](testing-and-troubleshooting.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f1e92-124">For information about how to resolve common issues that users may encounter with your Office Add-in, see [Troubleshoot user errors with Office Add-ins](testing-and-troubleshooting.md).</span></span>
