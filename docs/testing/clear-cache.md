---
title: Office のキャッシュをクリアする
description: コンピューターで Office のキャッシュをクリアする方法について説明します。
ms.date: 05/22/2020
localization_priority: Normal
ms.openlocfilehash: c48f3ed6f4c2f5f246341b6b878a725a54758bbe
ms.sourcegitcommit: 19312a54f47a17988ffa86359218a504713f9f09
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/10/2020
ms.locfileid: "44679403"
---
# <a name="clear-the-office-cache"></a>Office のキャッシュをクリアする

以前に Windows、Mac、または iOS にサイドロードしたアドインは、コンピューターで Office のキャッシュをクリアすることにより削除できます。

また、アドインのマニフェストに変更を加えた場合は (アイコンのファイル名やアドイン コマンドのテキストを更新した場合など)、Office のキャッシュをクリアし、更新されたマニフェストを使用してアドインをサイドロードし直す必要があります。 これを実行することにより、アドインは更新されたマニフェストの記載どおりに Office で表示されるようになります。

## <a name="clear-the-office-cache-on-windows"></a>Windows で Office のキャッシュをクリアする

Excel、Word、および PowerPoint からすべてのサイドロードアドインを削除するには、フォルダーの内容を削除します。

```text
%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\
```

サイドロードアドインを Outlook から削除するには、「[テスト用に outlook アドインをサイドロード](../outlook/sideload-outlook-add-ins-for-testing.md)する」に記載されている手順を使用して、インストールされているアドインを一覧表示するダイアログボックスの [**カスタムアドイン**] セクションで、アドインを検索します。アドインの省略記号 () を選択 `...` し、[**削除**] を選択してその特定のアドインを このアドインの削除が機能しない場合は、 `Wef` 前に説明したように、「Excel、Word、および PowerPoint」で説明したように、フォルダーの内容を削除します。

また、アドインが Microsoft Edge で実行されているときに Windows 10 で Office のキャッシュをクリアするには、Microsoft Edge DevTools を使用します。

> [!TIP]
> サイドロードアドインで、HTML または JavaScript ソースファイルへの最新の変更を反映させる場合は、キャッシュをクリアする必要はありません。 代わりに、アドインの作業ウィンドウにフォーカスを置き (タスク ウィンドウ内の任意の場所をクリック)、**F5** キーを押してアドインをリロードします。

> [!NOTE]
> 次の手順を使用して Office のキャッシュをクリアするには、アドインに作業ウィンドウが必要です。 アドインが UI を使用しない場合 (たとえば、[送信時](../outlook/outlook-on-send-addins.md)機能を使用するアドインの場合)、次の手順でキャッシュをクリアする前に、同じドメインを [SourceLocation](../reference/manifest/sourcelocation.md) に使用するアドインに作業ウィンドウを追加する必要があります。

1. [Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj) をインストールします。

2. アドインを Office クライアントで開きます。

3. Microsoft Edge DevTools を実行します。

4. Microsoft Edge DevTools で、[**ローカル**] タブを開きます。アドインの名前が一覧表示されます。

5. アドイン名を選択して、アドインにデバッガーをアタッチします。 デバッガーがアドインにアタッチされると、新しい Microsoft Edge DevTools ウィンドウが開きます。

6. 新しいウィンドウの [**ネットワーク**] タブで、[**キャッシュのクリア**] ボタンを選択します。

    ![[キャッシュのクリア] ボタンが強調表示された Microsoft Edge DevTools のスクリーンショット](../images/edge-devtools-clear-cache.png)

7. これらの手順を完了しても望む結果が得られない場合は、[**常にサーバーから更新する**] ボタンを選択することもできます。

    ![[常にサーバーから更新する] ボタンが強調表示された Microsoft Edge DevTools のスクリーンショット](../images/edge-devtools-refresh-from-server.png)

## <a name="clear-the-office-cache-on-mac"></a>Mac で Office のキャッシュをクリアする

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

## <a name="clear-the-office-cache-on-ios"></a>iOS で Office のキャッシュをクリアする

iOS で Office のキャッシュをクリアするには、アドイン内の JavaScript から `window.location.reload(true)` を呼び出し、強制的に再読み込みを行います。 別の方法として、Office を再インストールすることもできます。

## <a name="see-also"></a>関連項目

- [Office アドインをデバッグする](debug-add-ins-using-f12-developer-tools-on-windows-10.md)
- [ランタイム ログを使用してアドインをデバッグする](runtime-logging.md)
- [テスト用に Office アドインをサイドロードする](sideload-office-add-ins-for-testing.md)
- [Office アドインの XML マニフェスト](../develop/add-in-manifests.md)
- [Office アドインのマニフェストを検証する](troubleshoot-manifest.md)
