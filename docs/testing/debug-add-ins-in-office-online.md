---
title: Office on the web でアドインをデバッグする
description: Office on the web を使用してアドインをテストおよびデバッグする方法。
ms.date: 03/06/2022
ms.localizationpriority: medium
ms.openlocfilehash: b365be937058f818a97dd7a73176a56f76b36098
ms.sourcegitcommit: a32f5613d2bb44a8c812d7d407f106422a530f7a
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/14/2022
ms.locfileid: "67674626"
---
# <a name="debug-add-ins-in-office-on-the-web"></a>Office on the web でアドインをデバッグする

この記事では、Office on the webを使用してアドインをデバッグする方法について説明します。次の手法を使用します。

- Windows や Office デスクトップ クライアント&mdash;を実行していないコンピューターでアドインをデバッグする場合 (Mac または Linux で開発している場合など)。
- Visual Studio や Visual Studio Code などの IDE でデバッグできない場合、またはデバッグしない場合は、別のデバッグ プロセスとして使用します。

この記事では、デバッグする必要があるアドイン プロジェクトがあることを前提としています。 Web 上でデバッグを行うだけの場合は、Word のクイック スタートなど、特定の Office アプリケーションのクイック スタートのいずれかを使用して新しいプロジェクト [を](../quickstarts/word-quickstart.md)作成します。

## <a name="debug-your-add-in"></a>アドインのデバッグ

Word on the web を使用してアドインをデバッグするには: 

1. localhost でプロジェクトを実行し、Office on the web内のドキュメントにサイドロードします。 サイドローディングの詳細な手順については、「 [Web 上の Office アドインを手動でサイドロードする」を](sideload-office-add-ins-for-testing.md#manually-sideload-an-add-in-to-office-on-the-web)参照してください。

2. ブラウザーの開発者ツールを開きます。 これは通常、F12 キーを押すことによって行われます。 デバッガー ツールを開き、それを使用してブレークポイントを設定し、変数を監視します。 ブラウザーのツールの使用に関する詳細なヘルプについては、次のいずれかを参照してください。

   - [Firefox](https://firefox-source-docs.mozilla.org/devtools-user/index.html)
   - [Safari](https://support.apple.com/guide/safari/use-the-developer-tools-in-the-develop-menu-sfri20948/mac)
   - [Microsoft Edge (Chromium ベース)で開発者ツールを使用してアドインをデバッグする](debug-add-ins-using-devtools-edge-chromium.md)
   - [Edge レガシー用の開発者ツールを使用してアドインをデバッグする](debug-add-ins-using-devtools-edge-legacy.md)

   > [!NOTE]
   > Office on the web Internet Explorer では開きません。

## <a name="potential-issues"></a>潜在的な問題

デバッグ時に発生する可能性のある問題を次に示します。

- 表示される JavaScript エラーのいくつかは Office on the web に起因している可能性があります。

- ブラウザーに無効な証明書エラーが表示されることがありますが、このエラーはバイパスする必要があります。 これを行うプロセスは、ブラウザおよびこの変更を定期的に行うさまざまなブラウザの UI によって異なります。 詳細については、ブラウザーのヘルプを検索するか、オンラインで検索してください。 (たとえば、「Microsoft Edge の無効な証明書警告」を検索します。) ほとんどのブラウザーには、警告ページにリンクがあり、このリンクをクリックするとアドイン ページにアクセスされます。 たとえば、Microsoft Edge には「Web ページへ移動 (推奨しません)」というリンクがあります。 ただし、通常はアドインが再び読み込まれるたびに、このリンクを経由する必要があります。 継続的なバイパスについては、お勧めのヘルプを参照してください。

- コードにブレークポイントを設定した場合、Office on the webは保存できないことを示すエラーをスローする可能性があります。

## <a name="see-also"></a>関連項目

- [Office アドイン開発のベスト プラクティス](../concepts/add-in-development-best-practices.md)
- [Office アドインでのユーザー エラーのトラブルシューティング](testing-and-troubleshooting.md)
