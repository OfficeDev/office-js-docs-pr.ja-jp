---
title: サイドロード コマンドを使用した Sideload Office アドイン
description: ''
ms.date: 07/24/2018
ms.openlocfilehash: 1ab0277493f2899adb479c2f24b1635a881af3cc
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2018
ms.locfileid: "23944042"
---
# <a name="sideload-office-add-ins-for-testing-using-the-sideload-command"></a>**サイドロードコマンド** を使用したテストの Sideload Office アドイン
 >[!NOTE]
>[Npm 実行 sideload] メソッドだけが、Excel、Word、および PowerPoint のアドイン ウィンドウ上で実行されます。作成されたアドインのプロジェクトに対してのみ、 [**yo office** ツール](https://github.com/OfficeDev/generator-office) があると、 `sideload` 内のスクリプトを作成、 `scripts` 、package.json ファイルのセクション。(プロジェクトの以前のバージョンで作成された **yo office** も、このスクリプトはありません)。プロジェクトは、Visual Studio で作成されたか、sideload スクリプトがない、 [Office アドインをネットワーク共有から Sideload](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)に記載されているメソッドを使用して Windows に sideload することができます。
>
> Word、Excel、PowerPoint のアドインを Windows でテストしない場合は、以下のいずれかのトピックを参照してアドインをサイドロードします。
> 
> - [テスト用に Office Online で Office アドインをサイドロードする](sideload-office-add-ins-for-testing.md)
> - [テスト用に iPad と Mac で Office アドインをサイドロードする](sideload-an-office-add-in-on-ipad-and-mac.md)
> - [テスト用に Outlook アドインをサイドロードする](https://docs.microsoft.com/outlook/add-ins/sideload-outlook-add-ins-for-testing)

1. コマンド プロンプトを管理者として開きます。

2. ディレクトリをアドイン プロジェクト フォルダのルートに変更します。

3. 次のコマンドを実行して、ポート 3000 上のローカル Web サーバー インスタンスを起動して、アドイン プロジェクトを提供します。「**npm run start**」

4. 二番目のコマンド プロンプトを管理者として開きます。

5. ディレクトリをアドイン プロジェクト フォルダのルートに変更します。

6. 次のコマンドを実行して、ホスト アプリケーション（Excel、Wordなど）を起動し、アドインをホスト アプリケーションに登録します。「**npm run sideload**」

## <a name="see-also"></a>関連項目

- [マニフェストの問題を検証し、トラブルシューティングする](troubleshoot-manifest.md)
- [Office アドインを発行する](../publish/publish.md)