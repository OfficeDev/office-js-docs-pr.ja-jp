---
title: Office アドインのテストとデバッグ
description: ''
ms.date: 11/24/2017
localization_priority: Priority
ms.openlocfilehash: 7ffa281807ca1541f8ebcc5f722c1043db115509
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/23/2019
ms.locfileid: "29388837"
---
# <a name="test-and-debug-office-add-ins"></a><span data-ttu-id="71048-102">Office アドインのテストとデバッグ</span><span class="sxs-lookup"><span data-stu-id="71048-102">Test and debug Office Add-ins</span></span>

<span data-ttu-id="71048-103">このセクションでは、Office アドインのテスト、デバッグ、トラブルシューティングに関するガイダンスを示します。</span><span class="sxs-lookup"><span data-stu-id="71048-103">This section contains guidance about testing, debugging, and troubleshooting issues with Office Add-ins.</span></span>

## <a name="sideload-an-office-add-in-for-testing"></a><span data-ttu-id="71048-104">テスト用に Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="71048-104">Sideload an Office Add-in for testing</span></span>

<span data-ttu-id="71048-105">サイドロードを使用すると、最初にアドイン カタログに置かなくても、テスト用に Office アドインをインストールすることができます。</span><span class="sxs-lookup"><span data-stu-id="71048-105">You can use sideloading to install an Office Add-in for testing without having to first put it in an add-in catalog.</span></span> <span data-ttu-id="71048-106">アドインをサイドロードする手順は、プラットフォームによって異なり、場合によっては、製品によっても異なります。</span><span class="sxs-lookup"><span data-stu-id="71048-106">The procedure for sideloading an add-in varies by platform, and in some cases, by product as well.</span></span> <span data-ttu-id="71048-107">次のそれぞれの記事では、特定のプラットフォームまたは特定の製品の Office アドインをサイドロードする方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="71048-107">The following articles each describe how to sideload Office Add-ins on a specific platform or within a specific product:</span></span>

- [<span data-ttu-id="71048-108">Windows で Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="71048-108">Sideload Office Add-ins on Windows</span></span>](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)

- [<span data-ttu-id="71048-109">Office Online で Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="71048-109">Sideload Office Add-ins in Office Online</span></span>](sideload-office-add-ins-for-testing.md)

- [<span data-ttu-id="71048-110">iPad と Mac で Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="71048-110">Sideload Office Add-ins on iPad and Mac</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)

- [<span data-ttu-id="71048-111">テスト用に Outlook アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="71048-111">Sideload Outlook add-ins for testing</span></span>](https://docs.microsoft.com/outlook/add-ins/sideload-outlook-add-ins-for-testing)

## <a name="debug-an-office-add-in"></a><span data-ttu-id="71048-112">Office アドインのデバッグ</span><span class="sxs-lookup"><span data-stu-id="71048-112">Debug an Office Add-in</span></span>

<span data-ttu-id="71048-113">Office アドインをデバッグする手順も、プラットフォームによって異なります。</span><span class="sxs-lookup"><span data-stu-id="71048-113">The procedure for debugging an Office Add-in varies by platform as well.</span></span> <span data-ttu-id="71048-114">次のそれぞれの記事では、特定のプラットフォームで Office アドインをデバッグする方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="71048-114">Each of the following articles describes how to debug Office Add-ins on a specific platform:</span></span>

- [<span data-ttu-id="71048-115">(Windows で) 作業ウィンドウからデバッガーをアタッチする</span><span class="sxs-lookup"><span data-stu-id="71048-115">Attach a debugger from the task pane (on Windows)</span></span>](attach-debugger-from-task-pane.md)

- [<span data-ttu-id="71048-116">Windows 10 で F12 開発者ツールを使用してアドインをデバッグする</span><span class="sxs-lookup"><span data-stu-id="71048-116">Debug add-ins using F12 developer tools on Windows 10</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)

- [<span data-ttu-id="71048-117">Office Online でアドインをデバッグする</span><span class="sxs-lookup"><span data-stu-id="71048-117">Debug add-ins in Office Online</span></span>](debug-add-ins-in-office-online.md)

- [<span data-ttu-id="71048-118">iPad と Mac で Office アドインをデバッグする</span><span class="sxs-lookup"><span data-stu-id="71048-118">Debug Office Add-ins on iPad and Mac</span></span>](debug-office-add-ins-on-ipad-and-mac.md)

## <a name="validate-an-office-add-in-manifest"></a><span data-ttu-id="71048-119">Office アドイン マニフェストの検証</span><span class="sxs-lookup"><span data-stu-id="71048-119">Validate an Office Add-in manifest</span></span>

<span data-ttu-id="71048-120">Office アドインを記述するマニフェスト ファイルを検証し、マニフェスト ファイルの問題のトラブルシューティングを行う方法については、「[マニフェストの問題を検証し、トラブルシューティングを行う](troubleshoot-manifest.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="71048-120">For information about how to validate the manifest file that describes your Office Add-in and troubleshoot issues with the manifest file, see [Validate and troubleshoot issues with your manifest](troubleshoot-manifest.md).</span></span>

## <a name="troubleshoot-user-errors"></a><span data-ttu-id="71048-121">ユーザーのエラーのトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="71048-121">Troubleshoot user errors</span></span>

<span data-ttu-id="71048-122">よくある Office アドインの問題の解決方法については、「[Office アドインでのユーザー エラーのトラブルシューティング](testing-and-troubleshooting.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="71048-122">For information about how to resolve common issues that users may encounter with your Office Add-in, see [Troubleshoot user errors with Office Add-ins](testing-and-troubleshooting.md).</span></span>
