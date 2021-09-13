---
title: Office のキャッシュをクリアする
description: コンピューターで Office のキャッシュをクリアする方法について説明します。
ms.date: 08/02/2021
ms.localizationpriority: high
ms.openlocfilehash: 2a0c4f5b904e8f63a3a361f1339e87785bf40eea
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2021
ms.locfileid: "59154630"
---
# <a name="clear-the-office-cache"></a>Office のキャッシュをクリアする

以前に Windows、Mac、または iOS にサイドロードしたアドインを削除するには、コンピューターで Office のキャッシュをクリアする必要があります。

さらに、アドインのマニフェストに変更を加えた場合 (アイコンのファイル名やアドイン コマンドのテキストの更新など)、Office キャッシュをクリアしてから、更新されたマニフェストを使用してアドインを再サイドロードする必要があります。これにより、Office は更新されたマニフェストで説明されているとおりにアドインをレンダリングできます。

> [!NOTE]
> Excel、OneNote、PowerPoint、または Word on the web からサイドロードされたアドインを削除するには、「[テストのために Office on the web に Office アドインをサイドロードする: サイドロードされたアドインを削除する](sideload-office-add-ins-for-testing.md#remove-a-sideloaded-add-in)」を参照してください。

## <a name="clear-the-office-cache-on-windows"></a>Windows で Office のキャッシュをクリアする

Excel、Word、および PowerPoint からサイドロードされたすべてのアドインを削除するには、次のフォルダーのコンテンツを削除します。

```
%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\
```

次のフォルダーが存在する場合は、そのコンテンツも削除してください。

```
%userprofile%\AppData\Local\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\#!123\INetCache\
```

サイドロードされたアドインを Outlook から削除するには、「[テスト用に Outlook アドインをサイドロードする](../outlook/sideload-outlook-add-ins-for-testing.md)」の手順を使用して、インストールされているアドインが一覧表示されたダイアログ ボックスの **カスタム アドイン** セクションでアドインを検索します。アドインの省略記号 (`...`) を選択し、[**削除**] を選択して、そのアドインを削除します。 このアドインの削除が機能しない場合は、Excel、Word、PowerPoint について前に説明した `Wef` フォルダーのコンテンツを削除します。

また、アドインが Microsoft Edge で実行されているときに Windows 10 で Office のキャッシュをクリアするには、Microsoft Edge DevTools を使用します。

> [!TIP]
> サイドロードされたアドインに HTML や JavaScript のソース ファイルへの最近の変更を反映させたいだけの場合は、キャッシュをクリアする必要はありません。 代わりに、アドインの作業ウィンドウにフォーカスを置き (タスク ウィンドウ内の任意の場所をクリック)、**F5** キーを押してアドインをリロードします。

> [!NOTE]
> 次の手順を使用して Office のキャッシュをクリアするには、アドインに作業ウィンドウが必要です。 アドインが UI を使用しない場合 (たとえば、[送信時](../outlook/outlook-on-send-addins.md)機能を使用するアドインの場合)、次の手順でキャッシュをクリアする前に、同じドメインを [SourceLocation](../reference/manifest/sourcelocation.md) に使用するアドインに作業ウィンドウを追加する必要があります。

1. [Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj) をインストールします。

2. アドインを Office クライアントで開きます。

3. Microsoft Edge DevTools を実行します。

4. Microsoft Edge DevTools で、[**ローカル**] タブを開きます。アドインの名前が一覧表示されます。

5. アドイン名を選択して、アドインにデバッガーをアタッチします。 デバッガーがアドインにアタッチされると、新しい Microsoft Edge DevTools ウィンドウが開きます。

6. 新しいウィンドウの [**ネットワーク**] タブで、[**キャッシュのクリア**] を選択します。

    ![[キャッシュのクリア] ボタンが強調表示された Microsoft Edge DevTools のスクリーンショット。](../images/edge-devtools-clear-cache.png)

7. これらの手順を完了しても望む結果が得られない場合は、[**常にサーバーから更新する**] を選択してみてください。

    ![[常にサーバーから更新する] ボタンが強調表示された Microsoft Edge DevTools のスクリーンショット。](../images/edge-devtools-refresh-from-server.png)

## <a name="clear-the-office-cache-on-mac"></a>Mac で Office のキャッシュをクリアする

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

## <a name="clear-the-office-cache-on-ios"></a>iOS で Office のキャッシュをクリアする

iOS で Office のキャッシュをクリアするには、アドイン内の JavaScript から `window.location.reload(true)` を呼び出し、強制的に再読み込みを行います。 または、Office を再インストールします。

## <a name="see-also"></a>関連項目

- [Office アドインでの開発エラーのトラブルシューティング](troubleshoot-development-errors.md)
- [Office アドインをデバッグする](debug-add-ins-using-f12-developer-tools-on-windows-10.md)
- [ランタイム ログを使用してアドインをデバッグする](runtime-logging.md)
- [テスト用に Office アドインをサイドロードする](sideload-office-add-ins-for-testing.md)
- [Office アドインの XML マニフェスト](../develop/add-in-manifests.md)
- [Office アドインのマニフェストを検証する](troubleshoot-manifest.md)
