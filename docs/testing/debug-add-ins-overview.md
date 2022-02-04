---
title: Office アドインをデバッグする
description: 開発環境向けの Office アドインのデバッグ ガイダンスを見つける
ms.date: 01/27/2022
ms.localizationpriority: high
ms.openlocfilehash: 490d2d786bbd7e3169e7202dbbd70e81f9525e41
ms.sourcegitcommit: e837f966d7360ed11b3ff9363ff20380f7d0c45e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/28/2022
ms.locfileid: "62263059"
---
# <a name="overview-of-debugging-office-add-ins"></a>Office アドインのデバッグの概要

Office アドイン のデバッグは、基本的に Web アプリケーションのデバッグと同じです。 ただし、単一のツール セットがすべてのアドイン開発者に対して機能するわけではありません。 これは、アドインをさまざまなオペレーティング システムで開発し、クロス プラットフォームで実行できるためです。 この記事は、開発環境の詳細なデバッグ ガイダンスを見つけるのに役立ちます。

> [!TIP]
> この記事は、ブレーク ポイントの設定とコードのステップ スルーという狭義のデバッグに関するものです。 テストとトラブルシューティングのガイダンスについては、「[Office アドインのテスト](test-debug-office-add-ins.md)」、「[Office アドインの開発エラーのトラブルシューティングを行う](troubleshoot-development-errors.md)」から始めます。

> [!NOTE]
> サポートするすべてのプラットフォームでアドインを *テスト* する必要がありますが、開発用コンピューターとは異なる環境で *デバッグ* する必要があることはほとんどありません。 このため、この記事では "開発用コンピューター" と "開発用環境" を使用して、デバッグしている環境を参照します。 コードの問題が開発用コンピューター以外のプラットフォームでのみ発生し、それを解決するためにブレーク ポイントを設定するかコードをステップ スルーする必要がある場合、デバッグしている環境は文字通りの開発環境ではありません。

## <a name="server-side-or-client-side"></a>サーバー側ですか、それともクライアント側ですか。

Office アドインのサーバー側コードのデバッグは、Web アプリケーションのサーバー側のデバッグと同じです。 IDE またはその他のツールのデバッグ手順を参照してください。 以下は、最も一般的なツールの例です。

- [Visual Studio で ASP.NET または ASP.NET Core アプリをデバッグする](/visualstudio/debugger/how-to-enable-debugging-for-aspnet-applications)
- [Express のデバッグ](https://expressjs.com/en/guide/debugging.html)
- [Node.js デバッグ ガイド](https://nodejs.org/en/docs/guides/debugging-getting-started/)
- [VS Code での Node.js のデバッグ](https://code.visualstudio.com/docs/nodejs/nodejs-debugging)
- [Webpack デバッグ](https://webpack.js.org/contribute/debugging/)

この記事の残りの部分は、クライアント側の JavaScript (TypeScript からトランスパイルされる可能性があります) のデバッグのみに関係しています。

クライアント側のコードをデバッグするためのガイダンスを見つけるために、最初の変数は開発用コンピューターのオペレーティング システムです。

- [Windows](#debug-on-windows)
- [Mac](#debug-on-mac)
- [Linux または他の Unix バリアント](#debug-on-linux)

## <a name="debug-on-windows"></a>Windows でデバッグする

以下に、Windows でのデバッグに関する一般的なガイダンスを示します。 Excel の UI なしのカスタム関数と Outlook のイベントベースのアドインをデバッグするための特別な手順があります。 このセクションで後述する「[Windows の特殊なケース](#special-cases-in-windows)」を参照してください。 Windows でのデバッグは、IDE によって異なります。

- **Visual Studio**: 内部デバッガーを使用してデバッグします。 「[Visual Studio で Office アドインをデバッグする](../develop/debug-office-add-ins-in-visual-studio.md)」を参照してください。
- **Visual Studio Code**: Visual Studio Code 用の [アドイン デバッガー拡張機能を使用してデバッグします](debug-with-vs-extension.md)。
- **その他の IDE** (または IDE 内でデバッグしたくない場合): アドインが開発コンピューターで使用するブラウザー ランタイムに関連付けられている開発者ツールを使用します。 次のいずれかをご覧ください。

    - [Internet Explorer の開発者ツールを使用してアドインをデバッグする](debug-add-ins-using-f12-tools-ie.md)
    - [Edge レガシー用の開発者ツールを使用してアドインをデバッグする](debug-add-ins-using-devtools-edge-legacy.md)
    - [Microsoft Edge (Chromium ベース)で開発者ツールを使用してアドインをデバッグする](debug-add-ins-using-devtools-edge-chromium.md)

使用されているブラウザー ランタイムについては、「[Office アドインで使用されるブラウザ](../concepts/browsers-used-by-office-web-add-ins.md)」を参照してください。

> [!TIP]
> [!INCLUDE[Identify the webview through the add-in UI](../includes/identify-webview-in-ui.md)]

### <a name="special-cases-in-windows"></a>Windows の特殊なケース

Windows で UI なしのカスタム関数をデバッグするには、「[UI なしのカスタム関数のデバッグ](../excel/custom-functions-debugging.md)」を参照してください。

Outlook でイベント ベースのアドインをデバッグするには、「[イベント ベースの Outlook アドインをデバッグする](../outlook/debug-autolaunch.md)」を参照してください。 このプロセスには、Visual Studio Code が必要です。

## <a name="debug-on-mac"></a>Mac でデバッグする

以下に、Mac でのデバッグに関する一般的なガイダンスを示します。 Excel で UI なしのカスタム関数をデバッグするための特別な手順があります。 このセクションで後述する「[Mac の特殊なケース](#special-cases-in-mac)」を参照してください。

- Visual Studio Code を使用している場合は、[Visual Studio Code 用のアドイン デバッガー拡張機能](debug-with-vs-extension.md)を使用してデバッグします。
- その他の IDE の場合は、Safari Web Inspector を使用してください。 手順については、「[Mac Officeアドインのデバッグ](debug-office-add-ins-on-ipad-and-mac.md)」を参照してください。

### <a name="special-cases-in-mac"></a>Mac の特殊なケース

Mac で UI なしのカスタム関数をデバッグするには、「[UI なしのカスタム関数のデバッグ](../excel/custom-functions-debugging.md)」を参照してください。

## <a name="debug-on-linux"></a>Linux でのデバッグ

Office for Linux のデスクトップ バージョンはないため、テストとデバッグを行うには、[Web 上の Office にアドインをサイドロードする](sideload-office-add-ins-for-testing.md)必要があります。 デバッグについてのガイダンスは「[Office on the web でアドインをデバッグする](debug-add-ins-in-office-online.md)」で確認できます。

> [!NOTE]
> Linux コンピューターで Office アドインを開発することはお勧めしません。ただし、すべてのアドインのユーザーが Linux コンピューターから Web 上の Office を介してアドインにアクセスすることが確実な場合を除きます。

## <a name="debug-add-ins-in-staging-or-production"></a>ステージングまたは運用でのアドインのデバッグ

既にステージングまたは運用にあるアドインをデバッグするには、アドインの UI からデバッガーをアタッチします。 手順については、「[作業ウィンドウからデバッガーをアタッチする](attach-debugger-from-task-pane.md)」を参照してください。