---
title: sideload コマンドを使用して Office アドインをサイドロードする
description: ''
ms.date: 06/20/2019
localization_priority: Priority
ms.openlocfilehash: 38aa74963ca750d65e4be7bb17745a59eeed0c83
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/21/2019
ms.locfileid: "35126891"
---
# <a name="sideload-office-add-ins-for-testing-using-the-sideload-command"></a><span data-ttu-id="bea3e-102">sideload コマンドを使用してテスト用の Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="bea3e-102">Sideload Office Add-ins for testing using the sideload command</span></span>
 
> [!NOTE]
> <span data-ttu-id="bea3e-103">この記事で説明するサイドローディング手法は、以下の場合にのみ有効です。</span><span class="sxs-lookup"><span data-stu-id="bea3e-103">The sideloading technique described in this article is only valid for:</span></span>
> 
> - <span data-ttu-id="bea3e-104">Windows 上で実行される Excel、Word、および PowerPoint のアドイン</span><span class="sxs-lookup"><span data-stu-id="bea3e-104">Excel, Word, and PowerPoint add-ins that run on Windows</span></span>
> 
> - <span data-ttu-id="bea3e-105">[Office アドイン用の Yeoman ジェネレーター](https://github.com/OfficeDev/generator-office)で作成され、package.json ファイルの `scripts` セクションに `sideload` スクリプトがあるアドイン プロジェクト。</span><span class="sxs-lookup"><span data-stu-id="bea3e-105">Add-in projects that were created with the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) and that have a `sideload` script in the `scripts` section of the package.json file.</span></span> <span data-ttu-id="bea3e-106">(Office アドイン用の Yeoman ジェネレーターの古いバージョンで作成されたプロジェクトには、このスクリプトはありません。)</span><span class="sxs-lookup"><span data-stu-id="bea3e-106">(Projects that were created with older versions of the Yeoman generator for Office Add-ins will not have this script.)</span></span>
 
<span data-ttu-id="bea3e-107">Office アドイン用の Yeoman ジェネレーターが提供する `sideload` スクリプトを使用してアドインをサイドロードするには、以下の手順を実行します。</span><span class="sxs-lookup"><span data-stu-id="bea3e-107">To sideload your add-in by using the `sideload` script that the Yeoman generator for Office Add-ins provides, complete the following steps:</span></span>

1. <span data-ttu-id="bea3e-108">管理者としてコマンド プロンプトを開きます。</span><span class="sxs-lookup"><span data-stu-id="bea3e-108">Open a command prompt as an administrator.</span></span>

2. <span data-ttu-id="bea3e-109">ディレクトリをアドイン プロジェクト フォルダーのルートに変更します。</span><span class="sxs-lookup"><span data-stu-id="bea3e-109">Change directories to the root of your add-in project folder.</span></span>

3. <span data-ttu-id="bea3e-110">次のコマンドを実行して、アドイン プロジェクトを提供するためにポート 3000 でローカル Web サーバーインスタンスを起動します。`npm run start`</span><span class="sxs-lookup"><span data-stu-id="bea3e-110">Run the following command to start a local web server instance on port 3000 to serve up your add-in project: `npm run start`</span></span>

4. <span data-ttu-id="bea3e-111">管理者として 2 番目のコマンド プロンプトを開きます。</span><span class="sxs-lookup"><span data-stu-id="bea3e-111">Open a second command prompt as an administrator.</span></span>

5. <span data-ttu-id="bea3e-112">ディレクトリをアドイン プロジェクト フォルダーのルートに変更します。</span><span class="sxs-lookup"><span data-stu-id="bea3e-112">Change directories to the root of your add-in project folder.</span></span>

6. <span data-ttu-id="bea3e-113">次のコマンドを実行してホスト アプリケーション (Excel、Wordなど) を起動し、アドインをホスト アプリケーションに登録します。`npm run sideload`</span><span class="sxs-lookup"><span data-stu-id="bea3e-113">Run the following command to boot the host application (e.g. Excel, Word) and register your add-in in the host application: `npm run sideload`</span></span>

<span data-ttu-id="bea3e-114">アドイン プロジェクトが Visual Studio で作成された場合、またはサイドロード スクリプトがない場合は、「[ネットワーク共有からの Office アドインのサイドロード](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)」で説明されている方法で Windows にサイドロードできます。</span><span class="sxs-lookup"><span data-stu-id="bea3e-114">If your add-in project was created with Visual Studio or does not have the sideload script, you can sideload it on Windows by using the method described in [Sideload an Office Add-in from a network share](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).</span></span>

<span data-ttu-id="bea3e-115">Windows で Word、Excel、または PowerPoint アドインをテストしていない場合は、アドインのサイドロードについて、次のトピックのいずれかを参照してください。</span><span class="sxs-lookup"><span data-stu-id="bea3e-115">If you're not testing a Word, Excel, or PowerPoint add-in on Windows, see one of the following topics for information about sideloading your add-in:</span></span>
 
- [<span data-ttu-id="bea3e-116">テスト用に Office on the web で Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="bea3e-116">Sideload Office Add-ins in Office on the web for testing</span></span>](sideload-office-add-ins-for-testing.md)
- [<span data-ttu-id="bea3e-117">テスト用に iPad と Mac で Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="bea3e-117">Sideload Office Add-ins on iPad and Mac for testing</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)
- [<span data-ttu-id="bea3e-118">テスト用に Outlook アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="bea3e-118">Sideload Outlook add-ins for testing</span></span>](/outlook/add-ins/sideload-outlook-add-ins-for-testing)

## <a name="see-also"></a><span data-ttu-id="bea3e-119">関連項目</span><span class="sxs-lookup"><span data-stu-id="bea3e-119">See also</span></span>

- [<span data-ttu-id="bea3e-120">マニフェストの問題を検証し、トラブルシューティングする</span><span class="sxs-lookup"><span data-stu-id="bea3e-120">Validate and troubleshoot issues with your manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="bea3e-121">Office アドインを発行する</span><span class="sxs-lookup"><span data-stu-id="bea3e-121">Publish your Office Add-in</span></span>](../publish/publish.md)
