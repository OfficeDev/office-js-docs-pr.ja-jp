---
title: Office のキャッシュをクリアする
description: コンピューターで Office のキャッシュをクリアする方法について説明します。
ms.date: 01/21/2020
localization_priority: Priority
ms.openlocfilehash: 68e5c022671844ee44bf8ca8ac00bc5af6564bad
ms.sourcegitcommit: 43166612e9b4bf7a73312a572663c8696353dbc6
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/29/2020
ms.locfileid: "41580968"
---
# <a name="clear-the-office-cache"></a>Office のキャッシュをクリアする

以前に Windows、Mac、または iOS にサイドロードしたアドインは、コンピューターで Office のキャッシュをクリアすることにより削除できます。 

また、アドインのマニフェストに変更を加えた場合は (アイコンのファイル名やアドイン コマンドのテキストを更新した場合など)、Office のキャッシュをクリアし、更新されたマニフェストを使用してアドインをサイドロードし直す必要があります。 これを実行することにより、アドインは更新されたマニフェストの記載どおりに Office で表示されるようになります。

## <a name="clear-the-office-cache-on-windows"></a>Windows で Office のキャッシュをクリアする

### <a name="excel-word-and-powerpoint"></a>Excel、Word、および PowerPoint 

Windows で Excel、Word、および PowerPoint の Office のキャッシュをクリアするには、フォルダー `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` 内のコンテンツを削除します。

### <a name="outlook-windows-10"></a>Outlook (Windows 10)

アドインが Microsoft Edge で実行されているときに Windows 10 で Office のキャッシュをクリアするには、Microsoft Edge DevTools を使用します。

> [!TIP]
> サイドロードされたアドインに HTML や JavaScript のソース ファイルへの最近の変更を反映させたいなら、次の手順でキャッシュをクリアする必要はありません。 代わりに、アドインの作業ウィンドウにフォーカスを置き (タスク ウィンドウ内の任意の場所をクリック)、**F5** キーを押してアドインをリロードします。 

> [!NOTE]
> 次の手順を使用して Outlook のキャッシュをクリアするには、アドインに作業ウィンドウが必要です。 アドインが UI を使用しない場合 (たとえば、[送信時](/outlook/add-ins/outlook-on-send-addins)機能を使用するアドインの場合)、次の手順でキャッシュをクリアする前に、同じドメインを [SourceLocation](../reference/manifest/sourcelocation.md) に使用するアドインに作業ウィンドウを追加する必要があります。

1. [Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj) をインストールします。

2. Outlook でアドインを開きます。

3. Microsoft Edge DevTools を実行します。

4. Microsoft Edge DevTools で、[**ローカル**] タブを開きます。アドインの名前が一覧表示されます。

5. アドイン名を選択して、アドインにデバッガーをアタッチします。 デバッガーがアドインにアタッチされると、新しい Microsoft Edge DevTools ウィンドウが開きます。

6. 新しいウィンドウの [**ネットワーク**] タブで、[**キャッシュのクリア**] ボタンを選択します。

    ![[キャッシュのクリア] ボタンが強調表示された Microsoft Edge DevTools のスクリーンショット](../images/edge-devtools-clear-cache.png)

7. これらの手順を完了しても望む結果が得られない場合は、[**常にサーバーから更新する**] ボタンを選択することもできます。

    ![[常にサーバーから更新する] ボタンが強調表示された Microsoft Edge DevTools のスクリーンショット](../images/edge-devtools-refresh-from-server.png)

## <a name="clear-the-office-cache-on-mac"></a>Mac で Office のキャッシュをクリアする

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

##  <a name="clear-the-office-cache-on-ios"></a>iOS で Office のキャッシュをクリアする

iOS で Office のキャッシュをクリアするには、アドイン内の JavaScript から `window.location.reload(true)` を呼び出し、強制的に再読み込みを行います。 別の方法として、Office を再インストールすることもできます。

## <a name="see-also"></a>関連項目

- [Office アドインをデバッグする](debug-add-ins-using-f12-developer-tools-on-windows-10.md)
- [ランタイム ログを使用してアドインをデバッグする](runtime-logging.md)
- [テスト用に Office アドインをサイドロードする](sideload-office-add-ins-for-testing.md)
- [Office アドインの XML マニフェスト](../develop/add-in-manifests.md)
- [Office アドインのマニフェストを検証する](troubleshoot-manifest.md)

