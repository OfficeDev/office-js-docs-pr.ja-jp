---
title: Office のキャッシュをクリアする
description: コンピューターで Office のキャッシュをクリアする方法について説明します。
ms.date: 01/27/2022
ms.localizationpriority: high
ms.openlocfilehash: 2c2b22ececbf293578c9467269c4ad6779eb9aea
ms.sourcegitcommit: e837f966d7360ed11b3ff9363ff20380f7d0c45e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/28/2022
ms.locfileid: "62263094"
---
# <a name="clear-the-office-cache"></a>Office のキャッシュをクリアする

以前に Windows、Mac、または iOS にサイドロードしたアドインを削除するには、コンピューターで Office のキャッシュをクリアする必要があります。

さらに、アドインのマニフェストに変更を加えた場合 (アイコンのファイル名やアドイン コマンドのテキストの更新など)、Office キャッシュをクリアしてから、更新されたマニフェストを使用してアドインを再サイドロードする必要があります。これにより、Office は更新されたマニフェストで説明されているとおりにアドインをレンダリングできます。

> [!NOTE]
> Excel、OneNote、PowerPoint、または Word on the web からサイドロードされたアドインを削除するには、「[テストのために Office on the web に Office アドインをサイドロードする: サイドロードされたアドインを削除する](sideload-office-add-ins-for-testing.md#remove-a-sideloaded-add-in)」を参照してください。

## <a name="clear-the-office-cache-on-windows"></a>Windows で Office のキャッシュをクリアする

Windows コンピューターの Office キャッシュをクリアするには、自動、手動、およびMicrosoft Edge開発者ツールの使用という 3 つの方法があります。 メソッドについては、次のサブセクションで説明します。

### <a name="automatically"></a>自動的に

この方法は、アドイン開発用コンピューターに推奨されます。 Office on Windows バージョンが 2108 以降の場合、次の手順では、次に Office を再度開いたときに Office キャッシュをクリアするように構成します。

> [!NOTE]
> 自動的な方法は Outlook ではサポートされていません。

1. Outlook を除く任意の Office ホストのリボンから、**ファイル** > **オプション** > **信頼できるセキュリティ センター** > **セキュリティ センターの設定** > **信頼できるアドイン カタログ** に移動します。
1. 次回の Office の起動時にチェック ボックス **を選択し、以前に起動したすべての Web アドイン キャッシュ** をクリアします。

### <a name="manually"></a>手動

Excel、Word、PowerPoint の手動の方法は Outlook とは異なります。

#### <a name="manually-clear-the-cache-in-excel-word-and-powerpoint"></a>Excel、Word、PowerPoint でキャッシュを手動でクリアする

Excel、Word、PowerPoint からサイドロードされたすべてのアドインを削除するには、次のフォルダーの内容を削除します。

```
%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\
```

次のフォルダーが存在する場合は、そのコンテンツも削除してください。

```
%userprofile%\AppData\Local\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\#!123\INetCache\
```

#### <a name="manually-clear-the-cache-in-outlook"></a>Outlook でキャッシュを手動でクリアする

サイドロードされたアドインを Outlook から削除するには、「[テスト用に Outlook アドインをサイドロードする](../outlook/sideload-outlook-add-ins-for-testing.md)」の手順を使用して、インストールされているアドインが一覧表示されたダイアログ ボックスの「**カスタム アドイン**」セクションでアドインを検索します。アドインの省略記号 (`...`) を選択し、[**削除**] を選択して、そのアドインを削除します。このアドインの削除が機能しない場合は、前述したとおり Excel、Word、PowerPoint で `Wef` フォルダの内容を削除します。

### <a name="using-the-microsoft-edge-developer-tools"></a>Microsoft Edge開発者ツールの使用

アドインが Microsoft Edge で実行されているときにWindows 10の Office キャッシュをクリアするには、Microsoft Edge DevTools を使用できます。

> [!TIP]
> サイドロードされたアドインに HTML や JavaScript のソース ファイルへの最近の変更を反映させたいだけの場合は、キャッシュをクリアする必要はありません。 代わりに、アドインの作業ウィンドウにフォーカスを置き (タスク ウィンドウ内の任意の場所をクリック)、**Ctrl + F5** を押してアドインをリロードします。

> [!NOTE]
> 次の手順を使用して Office のキャッシュをクリアするには、アドインに作業ウィンドウが必要です。アドインが UI を使用しない場合 (たとえば、[送信時](../outlook/outlook-on-send-addins.md)機能を使用するアドインの場合)、次の手順でキャッシュをクリアする前に、同じドメインを [SourceLocation](../reference/manifest/sourcelocation.md) に使用するアドインに作業ウィンドウを追加する必要があります。

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

iOS で Office のキャッシュをクリアするには、アドイン内の JavaScript から `window.location.reload(true)` を呼び出し、強制的に再読み込みを行います。または、Office を再インストールします。

## <a name="see-also"></a>関連項目

- [Office アドインでの開発エラーのトラブルシューティング](troubleshoot-development-errors.md)
- [Internet Explorer の開発者ツールを使用してアドインをデバッグする](debug-add-ins-using-f12-tools-ie.md)
- [Edge レガシー用の開発者ツールを使用してアドインをデバッグする](debug-add-ins-using-devtools-edge-legacy.md)
- [Microsoft Edge (Chromium ベース)で開発者ツールを使用してアドインをデバッグする](debug-add-ins-using-devtools-edge-chromium.md)
- [ランタイム ログを使用してアドインをデバッグする](runtime-logging.md)
- [テスト用に Office アドインをサイドロードする](sideload-office-add-ins-for-testing.md)
- [Office アドインの XML マニフェスト](../develop/add-in-manifests.md)
- [Office アドインのマニフェストを検証する](troubleshoot-manifest.md)
