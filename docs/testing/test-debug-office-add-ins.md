---
title: Office アドインのテストとデバッグ
description: Office アドインのテストとデバッグを行う方法について説明します。
ms.date: 05/19/2021
localization_priority: Priority
ms.openlocfilehash: f794225d5ece20a430b967c8aa81ea165b573e52
ms.sourcegitcommit: 0d3bf72f8ddd1b287bf95f832b7ecb9d9fa62a24
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/02/2021
ms.locfileid: "52727928"
---
# <a name="test-and-debug-office-add-ins"></a>Office アドインのテストとデバッグ

この記事では、Office アドインのテスト、デバッグ、トラブルシューティングに関するガイダンスを示します。

## <a name="test-cross-platform-and-for-multiple-versions-of-office"></a>クロスプラットフォームおよび複数バージョンの Office をテストする

Office アドインは主要なプラットフォームで実行されるため、ユーザーが Office を実行している可能性のあるすべてのプラットフォームでアドインをテストする必要があります。 これには通常、Office on the web、Office on Windows (サブスクリプションと 1 回限りの購入の両方)、Office on Mac、Office on iOS、および (Outlook アドインの場合) Office on Android が含まれます。 ただし、一部のプラットフォームで作業しているユーザーがいないことを確認できる場合もあります。 たとえば、ユーザーが Windows コンピューターとサブスクリプション Office を使用する必要がある会社のアドインを作成する場合、Office on Mac や 1 回限りの購入の Windows をテストする必要はありません。 

> [!NOTE]
> Windows コンピューターでは、Windows と Office のバージョンによって、アドインが使用するブラウザー コントロールが決まります。詳細については、「[Office アドインによって使用されるブラウザー](../concepts/browsers-used-by-office-web-add-ins.md)」を参照してください。

> [!IMPORTANT]
> AppSource を通じて販売されるアドインは、すべてのプラットフォームでのテストを含む検証プロセスを経ます。 さらに、アドインは、Microsoft Edge (Chromium ベースの WebView2)、Chrome、Safari など、すべての主要な最新のブラウザーで Office on the web 用にテストされています。 したがって、AppSource に送信する前に、これらのプラットフォームとブラウザーでテストする必要があります。 検証の詳細については、「[コマーシャル マーケットプレースの認定ポリシー](/legal/marketplace/certification-policies)」、特に[セクション 1120.3](/legal/marketplace/certification-policies#11203-functionality)、および [Office アドイン アプリケーションと可用性のページ](../overview/office-add-in-availability.md)を参照してください。 
>
> AppSource は、Office on the web でアドインをテストするために、Internet Explorer または Microsoft Edge の以前のバージョン (WebView1) を使用しません。 ただし、かなりの数のユーザーがこれら 2 つのブラウザーを使用して Office on the web を開く場合は、それらのブラウザーでテストする必要があります。 詳細については、「[Internet Explorer 11 のサポート](../develop/support-ie-11.md)」および「[Microsoft Edge の問題のトラブルシューティング](../concepts/browsers-used-by-office-web-add-ins.md#troubleshooting-microsoft-edge-issues)」を参照してください。 Office は引き続きアドイン用にこれらのブラウザーをサポートしているため、アドインの実行時にバグが発生したと思われる場合は、[office-js](https://github.com/OfficeDev/office-js/issues/new/choose) リポジトリの問題を作成してください。

## <a name="sideload-an-office-add-in-for-testing"></a>テスト用に Office アドインをサイドロードする

サイドロードを使用すると、最初にアドイン カタログに置かなくても、テスト用に Office アドインをインストールすることができます。 アドインをサイドロードする手順は、プラットフォームによって異なり、場合によっては、製品によっても異なります。 次のそれぞれの記事では、特定のプラットフォームまたは特定の製品の Office アドインをサイドロードする方法について説明します。

- [Windows で Office アドインをサイドロードする](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)

- [Office on the web で Office アドインをサイドロードする](sideload-office-add-ins-for-testing.md)

- [iPad と Mac で Office アドインをサイドロードする](sideload-an-office-add-in-on-ipad-and-mac.md)

- [テスト用に Outlook アドインをサイドロードする](../outlook/sideload-outlook-add-ins-for-testing.md)

## <a name="debug-an-office-add-in"></a>Office アドインのデバッグ

Office アドインをデバッグする手順も、プラットフォームによって異なります。 次のそれぞれの記事では、特定のプラットフォームで Office アドインをデバッグする方法について説明します。

- [(Windows で) 作業ウィンドウからデバッガーをアタッチする](attach-debugger-from-task-pane.md)

- [Windows 10 で F12 開発者ツールを使用してアドインをデバッグする](debug-add-ins-using-f12-developer-tools-on-windows-10.md)

- [Office on the web でアドインをデバッグする](debug-add-ins-in-office-online.md)

- [iPad と Mac で Office アドインをデバッグする](debug-office-add-ins-on-ipad-and-mac.md)

- [Visual Studio Code 用 Microsoft Office アドイン デバッガー拡張機能](debug-with-vs-extension.md)

## <a name="validate-an-office-add-in-manifest"></a>Office アドイン マニフェストの検証

Office アドインを記述するマニフェスト ファイルを検証し、マニフェスト ファイルの問題のトラブルシューティングを行う方法については、「[マニフェストの問題を検証し、トラブルシューティングを行う](troubleshoot-manifest.md)」を参照してください。

## <a name="troubleshoot-user-errors"></a>ユーザーのエラーのトラブルシューティング

よくある Office アドインの問題の解決方法については、「[Office アドインでのユーザー エラーのトラブルシューティング](testing-and-troubleshooting.md)」を参照してください。
