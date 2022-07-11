---
title: Mac で Office アドインをデバッグする
description: Mac を使用して Office アドインをデバッグする方法について説明します。
ms.date: 03/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: 32d896743932abc7cf8be6bd62a491fc93fe0d1b
ms.sourcegitcommit: d8ea4b761f44d3227b7f2c73e52f0d2233bf22e2
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/11/2022
ms.locfileid: "66713001"
---
# <a name="debug-office-add-ins-on-a-mac"></a>Mac で Office アドインをデバッグする

アドインは HTML と JavaScript を使用して開発されているため、さまざまなプラットフォームで機能するように設計されていますが、さまざまなブラウザーで HTML の表示方法に微妙な違いがあります。この記事では、Mac で動作するアドインをデバッグする方法を説明します。

## <a name="debugging-with-safari-web-inspector-on-a-mac"></a>Mac での Safari Web インスペクタを使用したデバッグ

作業ウィンドウまたはコンテンツ アドインに UI を表示するアドインを使用している場合は、Safari Web インスペクタを使用して Office アドインをデバッグできます。

Mac で Office アドインをデバッグできるようにするには、Mac OS High Sierra AND Mac Office バージョン 16.9.1 (ビルド 18012504) 以降が必要です。 Office Mac ビルドがない場合は、 [Microsoft 365 開発者プログラム](https://developer.microsoft.com/office/dev-program)に参加することで入手できます。

最初に端末を開き、該当する Office アプリケーションの `OfficeWebAddinDeveloperExtras` プロパティを以下のように設定します。

- `defaults write com.microsoft.Word OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Excel OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Powerpoint OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Outlook OfficeWebAddinDeveloperExtras -bool true`

    > [!IMPORTANT]
    > Mac App Store Office のビルドでは、フラグは`OfficeWebAddinDeveloperExtras`サポートされていません。

次に Office アプリケーションを開き、[アドインをサイドロードします](sideload-an-office-add-in-on-mac.md)。 アドインを右クリックします。コンテキスト メニューに **[要素の検査]** オプションが表示されるはずです。 このオプションを選択するとインスペクタが表示されます。インスペクタでは、ブレークポイントを設定してアドインをデバッグできます。

> [!NOTE]
> インスペクタとダイアログ フリッカーを使おうとしている場合は、Office を最新バージョンに更新してください。 ちらつきが解決しない場合は、次の回避策を試してください。
>
> 1. ダイアログのサイズを変更します。
> 1. **[要素の検査]** を選択します (新しいウィンドウが開きます)。
> 1. ダイアログを元のサイズに変更します。
> 1. 必要に応じてインスペクタを使用します。

## <a name="clearing-the-office-applications-cache-on-a-mac"></a>Mac 上の Office アプリケーションのキャッシュをクリアする

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]
