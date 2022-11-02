---
title: Office アドインでのユーザー エラーのトラブルシューティング
description: Office アドインのユーザー エラーをトラブルシューティングする方法について説明します。
ms.date: 06/10/2022
ms.localizationpriority: medium
ms.openlocfilehash: 18bb3c180cd3af1eb8d045d7c69b9772532b04d4
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810373"
---
# <a name="troubleshoot-user-errors-with-office-add-ins"></a>Office アドインでのユーザー エラーのトラブルシューティング

At times your users might encounter issues with Office Add-ins that you develop. For example, an add-in fails to load or is inaccessible. Use the information in this article to help resolve common issues that your users encounter with your Office Add-in.

また、[Fiddler](https://www.telerik.com/fiddler) を使用して、アドインの問題を特定してデバッグすることもできます。

## <a name="common-errors-and-troubleshooting-steps"></a>一般的なエラーとトラブルシューティングの手順

次の表は、ユーザーが遭遇する可能性がある一般的なエラー メッセージとエラーを解決するためにユーザーが実行できる手順を示しています。

|**エラー メッセージ**|**解決策**|
|:-----|:-----|
|アプリのエラー: カタログに到達できませんでした|Verify firewall settings."Catalog" refers to AppSource. This message indicates that the user cannot access AppSource.|
|APP ERROR: This app could not be started. Close this dialog to ignore the problem or click "Restart" to try again.|Office の最新の更新プログラムがインストールされていることを確認するか、[Office 2013 更新プログラム ](https://support.microsoft.com/kb/2986156/)をダウンロードします。|
|エラー: オブジェクトがプロパティまたはメソッド 'defineProperty' をサポートしていません|Internet Explorerが互換モードで実行されていないことを確認します。 **[ツール** > **の互換性ビューの設定]** に移動します。|
|ブラウザーのバージョンがサポートされていないため、アプリを読み込めませんでした。 サポートされているブラウザーのバージョンの一覧についてはここをクリックしてください。|Make sure that the browser supports HTML5 local storage, or reset your Internet Explorer settings. For information about supported browsers, see [Requirements for running Office Add-ins](../concepts/requirements-for-running-office-add-ins.md).|

## <a name="when-installing-an-add-in-you-see-error-loading-add-in-in-the-status-bar"></a>アドインをインストールすると、ステータス バーに "アドイン読み込み中のエラー" と表示される

1. Office を終了します。
1. マニフェストが有効であることを確認します。 [「Office アドインのマニフェストを検証する](troubleshoot-manifest.md)」を参照してください。
1. アドインを再起動します。
1. もう一度アドインをインストールする。

また、フィードバックを寄せることができます。Windows または Mac 用 Excel を使用している場合は、Excel から直接 Office の機能拡張チームにフィードバックを送信できます。 これを行うには、[**ファイル**] > [**フィードバック**] > [**問題点、改善の報告**] の順に選択します。 問題点、改善の報告により、問題を理解するために必要なログが提供されます。

## <a name="outlook-add-in-doesnt-work-correctly"></a>Outlook アドインが正常に機能しない

Windows で実行され、[Internet Explorer を使用している](../concepts/browsers-used-by-office-web-add-ins.md) Outlook アドインが正常に機能しない場合は、Internet Explorer でスクリプトのデバッグを有効にしてみてください。

- **[ツール]** > **[インターネット オプション] [詳細設定]** >  の順に移動 **します**。
- **[参照]** で、**[スクリプトのデバッグを無効にする (Internet Explorer)]** と **[スクリプトのデバッグを無効にする (その他)]** の各チェックボックスをオフにします。

We recommend that you uncheck these settings only to troubleshoot the issue. If you leave them unchecked, you will get prompts when you browse. After the issue is resolved, check **Disable script debugging (Internet Explorer)** and **Disable script debugging (Other)** again.

## <a name="add-in-doesnt-activate-in-office-2013"></a>Office 2013 でアドインがアクティブにならない

ユーザーが次の手順を実行したときに、アドインがアクティブにならない場合があります。

1. Office 2013 で自分の Microsoft アカウントでサインインする。

1. 自分の Microsoft アカウントの 2 段階検証を有効にする。

1. アドインを挿入しようとする際に、メッセージに従って ID の確認を行う。

Office の最新の更新プログラムがインストールされていることを確認するか、[Office 2013 更新プログラム](https://support.microsoft.com/kb/2986156/)をダウンロードしてください。

## <a name="add-in-dialog-box-cannot-be-displayed"></a>アドイン ダイアログ ボックスを表示できない

When using an Office Add-in, the user is asked to allow a dialog box to be displayed. The user chooses **Allow**, and the following error message occurs.

"The security settings in your browser prevent us from creating a dialog box. Try a different browser, or configure your browser so that [URL] and the domain shown in your address bar are in the same security zone."

![ダイアログ ボックスのエラー メッセージのスクリーン ショット。](../images/dialog-prevented.png)

|影響を受けるブラウザー|影響を受けるプラットフォーム|
|:--------------------|:---------------------|
|Microsoft Edge|Office on the web|

この問題を解決するために、エンド ユーザーまたは管理者は、Microsoft Edge ブラウザー の信頼済みサイトのリストにアドインのドメインを追加することができます。

> [!IMPORTANT]
> アドインを信頼しない場合は、信頼済みサイトのリストにアドインの URL を追加しないでください。

URL を信頼済みサイトのリストに追加する方法:

1. [**コントロール パネル**]で、[**インターネット オプション**]  >  [**セキュリティ**] に移動します。
1. **[信頼済みサイト]** ゾーンを選択して、**[サイト]** を選択します。
1. エラー メッセージに表示される URL を入力して、**[追加]** を選択します。
1. Try to use the add-in again. If the problem persists, verify the settings for the other security zones and ensure that the add-in domain is in the same zone as the URL that is displayed in the address bar of the Office application.

This issue occurs when the Dialog API is used in pop-up mode. To prevent this issue from occurring, use the [displayInFrame](/javascript/api/office/office.ui) flag. This requires that your page support display within an iframe. The following example shows how to use the flag.

```js
Office.context.ui.displayDialogAsync(startAddress, {displayInIFrame:true}, callback);
```

## <a name="see-also"></a>関連項目

- [Office アドインでの開発エラーのトラブルシューティング](troubleshoot-development-errors.md)
