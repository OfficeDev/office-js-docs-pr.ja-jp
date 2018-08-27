---
title: サイドロード コマンドを使用した Sideload Office アドイン
description: ''
ms.date: 07/24/2018
ms.openlocfilehash: 3aacfdb09f362ea10ba0e2393caca335fe4c04c6
ms.sourcegitcommit: 4de2a1b62ccaa8e51982e95537fc9f52c0c5e687
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/10/2018
ms.locfileid: "22925102"
---
# <a name="sideload-office-add-ins-for-testing-using-the-sideload-command"></a><span data-ttu-id="c84e5-102">**サイドロードコマンド** を使用したテストの Sideload Office アドイン</span><span class="sxs-lookup"><span data-stu-id="c84e5-102">Sideload Office Add-ins for testing using the **sideload command**</span></span>
 >[!NOTE]
><span data-ttu-id="c84e5-103">「npm run sideload」メソッドは、Windows 上で実行されるExcel、Word、および PowerPoint アドインでのみ機能します。[**yo office** ツールで作成され](https://github.com/OfficeDev/generator-office)、 package.json ファイルの`sideload`  セクションに`scripts`  スクリプトのあるアドイン プロジェクトのみを対象とします。</span><span class="sxs-lookup"><span data-stu-id="c84e5-103">The "npm run sideload" method only works for Excel, Word, and PowerPoint add-ins that run on Windows; and only for add-in projects that were created with the [**yo office** tool](https://github.com/OfficeDev/generator-office) and that have a `sideload` script in the `scripts` section of the package.json file.</span></span> <span data-ttu-id="c84e5-104">(**yo office** の古いバージョンで作成されたプロジェクトには、このスクリプトもありません。) プロジェクトが Visual Studio で作成されている、またはサイドロードスクリプトがない場合は、[ネットワーク共有から Office アドインをサイドロードする](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)で記述したメソッドを使用して、Windows でサイドロードできます。</span><span class="sxs-lookup"><span data-stu-id="c84e5-104">(Projects that were created with older versions of **yo office** also do not have this script.) If your project was created with Visual Studio or does not have the sideload script, you can sideload it on Windows with the method described in [Sideload an Office add-in from a network share](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).</span></span>
>
> <span data-ttu-id="c84e5-105">Word、Excel、PowerPoint のアドインを Windows でテストしない場合は、以下のいずれかのトピックを参照してアドインをサイドロードします。</span><span class="sxs-lookup"><span data-stu-id="c84e5-105">If you're not testing a Word, Excel, or PowerPoint add-in on Windows, see one of the following topics to sideload your add-in:</span></span>
> 
> - [<span data-ttu-id="c84e5-106">テスト用に Office Online で Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="c84e5-106">Sideload Office Add-ins in Office Online for testing</span></span>](sideload-office-add-ins-for-testing.md)
> - [<span data-ttu-id="c84e5-107">テスト用に iPad と Mac で Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="c84e5-107">Sideload Office Add-ins on iPad and Mac for testing</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)
> - [<span data-ttu-id="c84e5-108">テスト用に Outlook アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="c84e5-108">Sideload Outlook Add-ins for testing</span></span>](https://docs.microsoft.com/outlook/add-ins/sideload-outlook-add-ins-for-testing)

1. <span data-ttu-id="c84e5-109">コマンド プロンプトを管理者として開きます。</span><span class="sxs-lookup"><span data-stu-id="c84e5-109">Open a command prompt as administrator:</span></span>

2. <span data-ttu-id="c84e5-110">ディレクトリをアドイン プロジェクト フォルダのルートに変更します。</span><span class="sxs-lookup"><span data-stu-id="c84e5-110">Change directories to the root of your add-in project folder.</span></span>

3. <span data-ttu-id="c84e5-111">次のコマンドを実行して、ポート 3000 上のローカル Web サーバー インスタンスを起動して、アドイン プロジェクトを提供します。「**npm run start**」</span><span class="sxs-lookup"><span data-stu-id="c84e5-111">Run the following command to start a local web server instance on port 3000 to serve up your add-in project: "**npm run start**"</span></span>

4. <span data-ttu-id="c84e5-112">二番目のコマンド プロンプトを管理者として開きます。</span><span class="sxs-lookup"><span data-stu-id="c84e5-112">Open a command prompt as administrator:</span></span>

5. <span data-ttu-id="c84e5-113">ディレクトリをアドイン プロジェクト フォルダのルートに変更します。</span><span class="sxs-lookup"><span data-stu-id="c84e5-113">Change directories to the root of your add-in project folder.</span></span>

6. <span data-ttu-id="c84e5-114">次のコマンドを実行して、ホスト アプリケーション（Excel、Wordなど）を起動し、アドインをホスト アプリケーションに登録します。「**npm run sideload**」</span><span class="sxs-lookup"><span data-stu-id="c84e5-114">Run the following command to boot the host application (e.g. Excel, Word) and register your add-in in the host application: "**npm run sideload**"</span></span>

## <a name="see-also"></a><span data-ttu-id="c84e5-115">関連項目</span><span class="sxs-lookup"><span data-stu-id="c84e5-115">See also</span></span>

- [<span data-ttu-id="c84e5-116">マニフェストの問題を検証し、トラブルシューティングする</span><span class="sxs-lookup"><span data-stu-id="c84e5-116">Validate and troubleshoot issues with your manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="c84e5-117">Office アドインを発行する</span><span class="sxs-lookup"><span data-stu-id="c84e5-117">Publish your Office Add-in</span></span>](../publish/publish.md)