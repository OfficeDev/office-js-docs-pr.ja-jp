---
title: サイドロード コマンドを使用した Sideload Office アドイン
description: ''
ms.date: 07/24/2018
ms.openlocfilehash: e831a1dfbc31ecf06c8b2d78dc1e9a8a4c9dcf01
ms.sourcegitcommit: 9e0952b3df852bd2896e9f4a6f59f5b89fc1ae24
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/27/2018
ms.locfileid: "21279361"
---
# <a name="sideload-office-add-ins-for-testing-using-the-sideload-command"></a><span data-ttu-id="a1bd8-102"> **サイドロードコマンド** を使用したテストの Sideload Office アドイン</span><span class="sxs-lookup"><span data-stu-id="a1bd8-102">Sideload Office Add-ins for testing using the **sideload command**</span></span>
 >[!NOTE]
><span data-ttu-id="a1bd8-103">「npm run sideload」メソッドは、Excel、Word、および PowerPoint アドインでのみ機能します）。</span><span class="sxs-lookup"><span data-stu-id="a1bd8-103">The "npm run sideload" method only works for Excel, Word, and PowerPoint add-ins).</span></span>

1. <span data-ttu-id="a1bd8-104">コマンド プロンプトを管理者として開きます。</span><span class="sxs-lookup"><span data-stu-id="a1bd8-104">Open a command prompt as administrator:</span></span>

2. <span data-ttu-id="a1bd8-105">ディレクトリをアドイン プロジェクト フォルダのルートに変更します。</span><span class="sxs-lookup"><span data-stu-id="a1bd8-105">Change directories to the root of your add-in project folder.</span></span>

3. <span data-ttu-id="a1bd8-106">次のコマンドを実行して、ポート 3000 上のローカル Web サーバー インスタンスを起動して、アドイン プロジェクトを提供します。「**npm run start**」</span><span class="sxs-lookup"><span data-stu-id="a1bd8-106">Run the following command to start a local web server instance on port 3000 to serve up your add-in project: "**npm run start**"</span></span>

4. <span data-ttu-id="a1bd8-107">コマンド プロンプトを管理者として開きます。</span><span class="sxs-lookup"><span data-stu-id="a1bd8-107">Open a command prompt as administrator:</span></span>

5. <span data-ttu-id="a1bd8-108">ディレクトリをアドイン プロジェクト フォルダのルートに変更します。</span><span class="sxs-lookup"><span data-stu-id="a1bd8-108">Change directories to the root of your add-in project folder.</span></span>

6. <span data-ttu-id="a1bd8-109">次のコマンドを実行して、ホスト アプリケーション（Excel、Wordなど）を起動し、アドインをホスト アプリケーションに登録します。「**npm run sideload**」</span><span class="sxs-lookup"><span data-stu-id="a1bd8-109">Run the following command to boot the host application (e.g. Excel, Word) and register your add-in in the host application: "**npm run sideload**"</span></span>

## <a name="see-also"></a><span data-ttu-id="a1bd8-110">関連項目</span><span class="sxs-lookup"><span data-stu-id="a1bd8-110">See also</span></span>

- [<span data-ttu-id="a1bd8-111">マニフェストの問題を検証し、トラブルシューティングする</span><span class="sxs-lookup"><span data-stu-id="a1bd8-111">Validate and troubleshoot issues with your manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="a1bd8-112">Office アドインを発行する</span><span class="sxs-lookup"><span data-stu-id="a1bd8-112">Publish your Office Add-in</span></span>](../publish/publish.md)