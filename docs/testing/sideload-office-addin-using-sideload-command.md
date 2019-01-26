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
# <a name="sideload-office-add-ins-for-testing-using-the-sideload-command"></a>**sideload コマンド**を使用して Office アドインをテストのためにサイドロードする
 >[!NOTE]
>"npm run sideload" メソッドは、Windows で実行される Excel、Word、および PowerPoint アドイン、[**yo office** ツール](https://github.com/OfficeDev/generator-office)を使って作成されたアドイン プロジェクト、および package.json ファイルの `scripts` セクションに `sideload` スクリプトが含まれているアドイン プロジェクトでのみ機能します。 (古いバージョンの **yo office** を使用して作成されたプロジェクトにはこのスクリプトがありません。) Visual Studio を使用して作成されたプロジェクト、または sideload スクリプトのないプロジェクトの場合、Windows でこれらのプロジェクトをサイドロードするには、「[テスト用に Office アドインをサイドロードする](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)」で説明するメソッドを使用します。
>
> Word、Excel、PowerPoint のアドインを Windows でテストしない場合は、以下のいずれかのトピックを参照してアドインをサイドロードします。
> 
> - [テスト用に Office Online で Office アドインをサイドロードする](sideload-office-add-ins-for-testing.md)
> - [テスト用に iPad と Mac で Office アドインをサイドロードする](sideload-an-office-add-in-on-ipad-and-mac.md)
> - [テスト用に Outlook アドインをサイドロードする](https://docs.microsoft.com/outlook/add-ins/sideload-outlook-add-ins-for-testing)

1. 管理者としてコマンド プロンプトを開きます。

2. ディレクトリをアドイン プロジェクト フォルダーのルートに変更します。

3. アドイン プロジェクトを処理するため、"**npm run start**" コマンドを実行してポート 3000 でローカル Web サーバーのインスタンスを開始します。

4. 管理者として 2 番目のコマンド プロンプトを開きます。

5. ディレクトリをアドイン プロジェクト フォルダーのルートに変更します。

6. "**npm run sideload**" コマンドを実行してホスト アプリケーション (Excel、Word など) を起動し、アドインをホスト アプリケーションに登録します。

## <a name="see-also"></a>関連項目

- [マニフェストの問題を検証し、トラブルシューティングする](troubleshoot-manifest.md)
- [Office アドインを発行する](../publish/publish.md)
