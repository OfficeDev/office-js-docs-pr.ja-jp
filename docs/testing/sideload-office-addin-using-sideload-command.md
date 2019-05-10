---
title: sideload コマンドを使用して Office アドインをサイドロードする
description: ''
ms.date: 03/19/201907/24/2018
localization_priority: Priority
ms.openlocfilehash: 69d39c2736312653b5a362aefccd41629e6e3555
ms.sourcegitcommit: 47b792755e655043d3db2f1fdb9a1eeb7453c636
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/06/2019
ms.locfileid: "33619078"
---
# <a name="sideload-office-add-ins-for-testing-using-the-sideload-command"></a>sideload コマンドを使用してテスト用の Office アドインをサイドロードする
 
> [!NOTE]
> この記事で説明するサイドローディング手法は、以下の場合にのみ有効です。
> 
> - Windows 上で実行される Excel、Word、および PowerPoint のアドイン
> 
> - [Office アドイン用の Yeoman ジェネレーター](https://github.com/OfficeDev/generator-office)で作成され、package.json ファイルの `scripts` セクションに `sideload` スクリプトがあるアドイン プロジェクト。 (Office アドイン用の Yeoman ジェネレーターの古いバージョンで作成されたプロジェクトには、このスクリプトはありません。)
 
Office アドイン用の Yeoman ジェネレーターが提供する `sideload` スクリプトを使用してアドインをサイドロードするには、以下の手順を実行します。

1. 管理者としてコマンド プロンプトを開きます。

2. ディレクトリをアドイン プロジェクト フォルダーのルートに変更します。

3. 次のコマンドを実行して、アドイン プロジェクトを提供するためにポート 3000 でローカル Web サーバーインスタンスを起動します。`npm run start`

4. 管理者として 2 番目のコマンド プロンプトを開きます。

5. ディレクトリをアドイン プロジェクト フォルダーのルートに変更します。

6. 次のコマンドを実行してホスト アプリケーション (Excel、Wordなど) を起動し、アドインをホスト アプリケーションに登録します。`npm run sideload`

アドイン プロジェクトが Visual Studio で作成された場合、またはサイドロード スクリプトがない場合は、「[ネットワーク共有からの Office アドインのサイドロード](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)」で説明されている方法で Windows にサイドロードできます。

Windows で Word、Excel、または PowerPoint アドインをテストしていない場合は、アドインのサイドロードについて、次のトピックのいずれかを参照してください。
 
- [テスト用に Office Online で Office アドインをサイドロードする](sideload-office-add-ins-for-testing.md)
- [テスト用に iPad と Mac で Office アドインをサイドロードする](sideload-an-office-add-in-on-ipad-and-mac.md)
- [テスト用に Outlook アドインをサイドロードする](/outlook/add-ins/sideload-outlook-add-ins-for-testing)

## <a name="see-also"></a>関連項目

- [マニフェストの問題を検証し、トラブルシューティングする](troubleshoot-manifest.md)
- [Office アドインを発行する](../publish/publish.md)
