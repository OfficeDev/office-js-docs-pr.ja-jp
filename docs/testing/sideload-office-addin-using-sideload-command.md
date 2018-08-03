---
title: サイドロード コマンドを使用した Sideload Office アドイン
description: ''
ms.date: 07/24/2018
ms.openlocfilehash: 90084fad0e79ab8acdf59eaa305825737401c0c8
ms.sourcegitcommit: e094aaa06d9aff3d13f8ffd3429d4a31f0b65b81
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/03/2018
ms.locfileid: "21782827"
---
# <a name="sideload-office-add-ins-for-testing-using-the-sideload-command"></a>**サイドロードコマンド** を使用したテストの Sideload Office アドイン
 >[!NOTE]
>「npm run sideload」メソッドは、Windows 上で実行されるExcel、Word、および PowerPoint アドインでのみ機能します。[**yo office** ツールで作成され](https://github.com/OfficeDev/generator-office)、 package.json ファイルの`sideload`  セクションに`scripts`  スクリプトのあるアドイン プロジェクトのみを対象とします。 (**yo office** の古いバージョンで作成されたプロジェクトには、このスクリプトもありません。) プロジェクトが Visual Studio で作成されている、またはサイドロードスクリプトがない場合は、[ネットワーク共有から Office アドインをサイドロードする](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)で記述したメソッドを使用して、Windows でサイドロードできます。
>
> Word、Excel、PowerPoint のアドインを Windows でテストしない場合は、以下のいずれかのトピックを参照してアドインをサイドロードします。
> 
> - [テスト用に Office Online で Office アドインをサイドロードする](sideload-office-add-ins-for-testing.md)
> - [テスト用に iPad と Mac で Office アドインをサイドロードする](sideload-an-office-add-in-on-ipad-and-mac.md)
> - [テスト用に Outlook アドインをサイドロードする](../../../../outlook/add-ins/sideload-outlook-add-ins-for-testing)

1. コマンド プロンプトを管理者として開きます。

2. ディレクトリをアドイン プロジェクト フォルダのルートに変更します。

3. 次のコマンドを実行して、ポート 3000 上のローカル Web サーバー インスタンスを起動して、アドイン プロジェクトを提供します。「**npm run start**」

4. 二番目のコマンド プロンプトを管理者として開きます。

5. ディレクトリをアドイン プロジェクト フォルダのルートに変更します。

6. 次のコマンドを実行して、ホスト アプリケーション（Excel、Wordなど）を起動し、アドインをホスト アプリケーションに登録します。「**npm run sideload**」

## <a name="see-also"></a>関連項目

- [マニフェストの問題を検証し、トラブルシューティングする](troubleshoot-manifest.md)
- [Office アドインを発行する](../publish/publish.md)