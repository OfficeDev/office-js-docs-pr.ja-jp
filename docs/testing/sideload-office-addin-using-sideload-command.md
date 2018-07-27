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
# <a name="sideload-office-add-ins-for-testing-using-the-sideload-command"></a> **サイドロードコマンド** を使用したテストの Sideload Office アドイン
 >[!NOTE]
>「npm run sideload」メソッドは、Excel、Word、および PowerPoint アドインでのみ機能します）。

1. コマンド プロンプトを管理者として開きます。

2. ディレクトリをアドイン プロジェクト フォルダのルートに変更します。

3. 次のコマンドを実行して、ポート 3000 上のローカル Web サーバー インスタンスを起動して、アドイン プロジェクトを提供します。「**npm run start**」

4. コマンド プロンプトを管理者として開きます。

5. ディレクトリをアドイン プロジェクト フォルダのルートに変更します。

6. 次のコマンドを実行して、ホスト アプリケーション（Excel、Wordなど）を起動し、アドインをホスト アプリケーションに登録します。「**npm run sideload**」

## <a name="see-also"></a>関連項目

- [マニフェストの問題を検証し、トラブルシューティングする](troubleshoot-manifest.md)
- [Office アドインを発行する](../publish/publish.md)