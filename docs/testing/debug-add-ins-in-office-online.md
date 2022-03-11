---
title: Office on the web でアドインをデバッグする
description: Office on the web を使用してアドインをテストおよびデバッグする方法。
ms.date: 03/06/2022
ms.localizationpriority: medium
ms.openlocfilehash: 5a07185c064d65432c7a3afce1e9f32e99034c3e
ms.sourcegitcommit: 3d7792b1f042db589edb74a895fcf6d7ced63903
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/11/2022
ms.locfileid: "63435691"
---
# <a name="debug-add-ins-in-office-on-the-web"></a>Office on the web でアドインをデバッグする

この記事では、アドインをOffice on the webする方法について説明します。次の方法を使用します。

- たとえば、Mac または Linux で開発している場合に、Windows または Office&mdash; デスクトップ クライアントを実行しないコンピューターでアドインをデバッグするには
- IDE でデバッグできない、または実行しない場合は、別のデバッグ プロセスとして、Visual StudioやVisual Studio Code。

この記事では、デバッグが必要なアドイン プロジェクトが含まれると想定しています。 Web でのデバッグの練習を行う場合は、Word のこのクイック スタートなど、特定の Office アプリケーションのクイック スタートのいずれかを使用して新しいプロジェクトを作成[します](../quickstarts/word-quickstart.md)。

## <a name="debug-your-add-in"></a>アドインのデバッグ

Word on the web を使用してアドインをデバッグするには: 

1. localhost でプロジェクトを実行し、プロジェクトをローカル ホストのドキュメントにサイドOffice on the web。 サイドローディングの手順の詳細については、「[Sideload Office Web 上のアドイン」を参照してください](sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web-manually)。

2. ブラウザーの開発者ツールを開きます。 これは通常、F12 キーを押して行います。 デバッガー ツールを開き、ブレークポイントとウォッチ変数の設定に使用します。 ブラウザーのツールの使用に関する詳細なヘルプについては、次のいずれかを参照してください。  

   - [Firefox](https://developer.mozilla.org/en-US/docs/Tools)
   - [Safari](https://support.apple.com/guide/safari/use-the-developer-tools-in-the-develop-menu-sfri20948/mac)
   - [Microsoft Edge (Chromium ベース)で開発者ツールを使用してアドインをデバッグする](debug-add-ins-using-devtools-edge-chromium.md)
   - [Edge レガシー用の開発者ツールを使用してアドインをデバッグする](debug-add-ins-using-devtools-edge-legacy.md)

   > [!NOTE]
   > Office on the webが開かInternet Explorer。

## <a name="potential-issues"></a>潜在的な問題

デバッグ時に発生する可能性があるいくつかの問題を次に示します。

- 表示される JavaScript エラーのいくつかは Office on the web に起因している可能性があります。

- ブラウザーに無効な証明書エラーが表示されることがありますが、このエラーはバイパスする必要があります。 これを行うプロセスは、ブラウザおよびこの変更を定期的に行うさまざまなブラウザの UI によって異なります。 詳細については、ブラウザーのヘルプを検索するか、オンラインで検索してください。 (たとえば、「Microsoft Edge の無効な証明書警告」を検索します。) ほとんどのブラウザーには、警告ページにリンクがあり、このリンクをクリックするとアドイン ページにアクセスされます。 たとえば、Microsoft Edge には「Web ページへ移動 (推奨しません)」というリンクがあります。 ただし、通常はアドインが再び読み込まれるたびに、このリンクを経由する必要があります。 継続的なバイパスについては、お勧めのヘルプを参照してください。

- コードにブレークポイントを設定すると、Office on the web保存できないというエラーがスローされる場合があります。

## <a name="see-also"></a>関連項目

- [Office アドイン開発のベスト プラクティス](../concepts/add-in-development-best-practices.md)
- [Office アドインでのユーザー エラーのトラブルシューティング](testing-and-troubleshooting.md)
