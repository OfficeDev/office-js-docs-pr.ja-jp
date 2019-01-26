---
title: sideload コマンドを使用して Office アドインをサイドロードする
description: ''
ms.date: 07/24/2018
localization_priority: Priority
ms.openlocfilehash: 2231e05d798dce4f4b5428627a3653ddcdecfc65
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/23/2019
ms.locfileid: "29387675"
---
# <a name="sideload-office-add-ins-for-testing-using-the-sideload-command"></a><span data-ttu-id="01ed6-102">**sideload コマンド**を使用して Office アドインをテストのためにサイドロードする</span><span class="sxs-lookup"><span data-stu-id="01ed6-102">Sideload Office Add-ins for testing using the **sideload command**</span></span>
 >[!NOTE]
><span data-ttu-id="01ed6-103">"npm run sideload" メソッドは、Windows で実行される Excel、Word、および PowerPoint アドイン、[**yo office** ツール](https://github.com/OfficeDev/generator-office)を使って作成されたアドイン プロジェクト、および package.json ファイルの `scripts` セクションに `sideload` スクリプトが含まれているアドイン プロジェクトでのみ機能します。</span><span class="sxs-lookup"><span data-stu-id="01ed6-103">The "npm run sideload" method only works for Excel, Word, and PowerPoint add-ins that run on Windows; and only for add-in projects that were created with the [**yo office** tool](https://github.com/OfficeDev/generator-office) and that have a `sideload` script in the `scripts` section of the package.json file.</span></span> <span data-ttu-id="01ed6-104">(古いバージョンの **yo office** を使用して作成されたプロジェクトにはこのスクリプトがありません。) Visual Studio を使用して作成されたプロジェクト、または sideload スクリプトのないプロジェクトの場合、Windows でこれらのプロジェクトをサイドロードするには、「[テスト用に Office アドインをサイドロードする](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)」で説明するメソッドを使用します。</span><span class="sxs-lookup"><span data-stu-id="01ed6-104">(Projects that were created with older versions of **yo office** also do not have this script.) If your project was created with Visual Studio or does not have the sideload script, you can sideload it on Windows with the method described in [Sideload an Office Add-in from a network share](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).</span></span>
>
> <span data-ttu-id="01ed6-105">Word、Excel、PowerPoint のアドインを Windows でテストしない場合は、以下のいずれかのトピックを参照してアドインをサイドロードします。</span><span class="sxs-lookup"><span data-stu-id="01ed6-105">If you're not testing a Word, Excel, or PowerPoint add-in on Windows, see one of the following topics to sideload your add-in:</span></span>
> 
> - [<span data-ttu-id="01ed6-106">テスト用に Office Online で Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="01ed6-106">Sideload Office Add-ins in Office Online for testing</span></span>](sideload-office-add-ins-for-testing.md)
> - [<span data-ttu-id="01ed6-107">テスト用に iPad と Mac で Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="01ed6-107">Sideload Office Add-ins on iPad and Mac for testing</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)
> - [<span data-ttu-id="01ed6-108">テスト用に Outlook アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="01ed6-108">Sideload Outlook add-ins for testing</span></span>](https://docs.microsoft.com/outlook/add-ins/sideload-outlook-add-ins-for-testing)

1. <span data-ttu-id="01ed6-109">管理者としてコマンド プロンプトを開きます。</span><span class="sxs-lookup"><span data-stu-id="01ed6-109">Open a command prompt as an administrator.</span></span>

2. <span data-ttu-id="01ed6-110">ディレクトリをアドイン プロジェクト フォルダーのルートに変更します。</span><span class="sxs-lookup"><span data-stu-id="01ed6-110">Change directories to the root of your add-in project folder.</span></span>

3. <span data-ttu-id="01ed6-111">アドイン プロジェクトを処理するため、"**npm run start**" コマンドを実行してポート 3000 でローカル Web サーバーのインスタンスを開始します。</span><span class="sxs-lookup"><span data-stu-id="01ed6-111">Run the following command to start a local web server instance on port 3000 to serve up your add-in project: "**npm run start**"</span></span>

4. <span data-ttu-id="01ed6-112">管理者として 2 番目のコマンド プロンプトを開きます。</span><span class="sxs-lookup"><span data-stu-id="01ed6-112">Open a second command prompt as an administrator.</span></span>

5. <span data-ttu-id="01ed6-113">ディレクトリをアドイン プロジェクト フォルダーのルートに変更します。</span><span class="sxs-lookup"><span data-stu-id="01ed6-113">Change directories to the root of your add-in project folder.</span></span>

6. <span data-ttu-id="01ed6-114">"**npm run sideload**" コマンドを実行してホスト アプリケーション (Excel、Word など) を起動し、アドインをホスト アプリケーションに登録します。</span><span class="sxs-lookup"><span data-stu-id="01ed6-114">Run the following command to boot the host application (e.g. Excel, Word) and register your add-in in the host application: "**npm run sideload**"</span></span>

## <a name="see-also"></a><span data-ttu-id="01ed6-115">関連項目</span><span class="sxs-lookup"><span data-stu-id="01ed6-115">See also</span></span>

- [<span data-ttu-id="01ed6-116">マニフェストの問題を検証し、トラブルシューティングする</span><span class="sxs-lookup"><span data-stu-id="01ed6-116">Validate and troubleshoot issues with your manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="01ed6-117">Office アドインを発行する</span><span class="sxs-lookup"><span data-stu-id="01ed6-117">Publish your Office Add-in</span></span>](../publish/publish.md)
