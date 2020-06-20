---
title: Office アドインでのユーザー エラーのトラブルシューティング
description: Office アドインでのユーザーエラーのトラブルシューティング方法について説明します。
ms.date: 06/17/2020
localization_priority: Normal
ms.openlocfilehash: 1dbc8cc18e0c9b12ccff605b655dd7c8629fb9cf
ms.sourcegitcommit: b939312ffdeb6e0a0dfe085db7efe0ff143ef873
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/19/2020
ms.locfileid: "44810850"
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
|エラー: オブジェクトがプロパティまたはメソッド 'defineProperty' をサポートしていません|Internet Explorerが互換モードで実行されていないことを確認します。 [ツール] > **[互換表示設定] ** に移動します。|
|Sorry, we couldn't load the app because your browser version is not supported. Click here for a list of supported browser versions.|Make sure that the browser supports HTML5 local storage, or reset your Internet Explorer settings. For information about supported browsers, see [Requirements for running Office Add-ins](../concepts/requirements-for-running-office-add-ins.md).|

## <a name="when-installing-an-add-in-you-see-error-loading-add-in-in-the-status-bar"></a>アドインをインストールすると、ステータス バーに "アドイン読み込み中のエラー" と表示される

1. Office を終了します。
2. マニフェストが有効であることを確認する
3. アドインを再起動する
4. もう一度アドインをインストールする。

また、フィードバックを寄せることができます。Windows または Mac 用 Excel を使用している場合は、Excel から直接 Office の機能拡張チームにフィードバックを送信できます。 これを行うには、[**ファイル**] | [**フィードバック**] | [**問題点、改善の報告**] の順に選択します。 問題点、改善の報告により、問題を理解するために必要なログが提供されます。

## <a name="outlook-add-in-doesnt-work-correctly"></a>Outlook アドインが正常に機能しない

Windows で実行され、[Internet Explorer を使用している](../concepts/browsers-used-by-office-web-add-ins.md) Outlook アドインが正常に機能しない場合は、Internet Explorer でスクリプトのデバッグを有効にしてみてください。 


- [ツール] > [**インターネットオプション**の詳細] に移動  >  **Advanced**します。
    
- **[参照]** で、**[スクリプトのデバッグを無効にする (Internet Explorer)]** と **[スクリプトのデバッグを無効にする (その他)]** の各チェックボックスをオフにします。
    
We recommend that you uncheck these settings only to troubleshoot the issue. If you leave them unchecked, you will get prompts when you browse. After the issue is resolved, check **Disable script debugging (Internet Explorer)** and **Disable script debugging (Other)** again.


## <a name="add-in-doesnt-activate-in-office-2013"></a>Office 2013 でアドインがアクティブにならない

ユーザーが次の手順を実行したときに、アドインがアクティブにならない場合があります。


1. Office 2013 で自分の Microsoft アカウントでサインインする。
    
2. 自分の Microsoft アカウントの 2 段階検証を有効にする。
    
3. アドインを挿入しようとする際に、メッセージに従って ID の確認を行う。
    
Office の最新の更新プログラムがインストールされていることを確認するか、[Office 2013 更新プログラム](https://support.microsoft.com/kb/2986156/)をダウンロードしてください。


## <a name="add-in-doesnt-load-in-task-pane-or-other-issues-with-the-add-in-manifest"></a>アドインが作業ウィンドウで読み込まれない、または他のアドイン マニフェストの問題

アドインのマニフェストでの問題をデバッグするには、「[Office アドインのマニフェストを検証する](troubleshoot-manifest.md)」および「[ランタイム ログを使用してアドインをデバッグする](runtime-logging.md)」を参照してください。


## <a name="add-in-dialog-box-cannot-be-displayed"></a>アドイン ダイアログ ボックスを表示できない

When using an Office Add-in, the user is asked to allow a dialog box to be displayed. The user chooses **Allow**, and the following error message occurs:

"The security settings in your browser prevent us from creating a dialog box. Try a different browser, or configure your browser so that [URL] and the domain shown in your address bar are in the same security zone."

![ダイアログ ボックスのエラー メッセージのスクリーン ショット](http://i.imgur.com/3mqmlgE.png)

|**影響を受けるブラウザー**|**影響を受けるプラットフォーム**|
|:--------------------|:---------------------|
|Internet Explorer、Microsoft Edge|Office on the web|

To resolve the issue, end users or administrators can add the domain of the add-in to the list of trusted sites in Internet Explorer. Use the same procedure whether you're using the Internet Explorer or Microsoft Edge browser.

> [!IMPORTANT]
> アドインを信頼しない場合は、信頼済みサイトのリストにアドインの URL を追加しないでください。

URL を信頼済みサイトのリストに追加する方法:

1. [**コントロール パネル**]で、[**インターネット オプション**]  >  [**セキュリティ**] に移動します。
2. **[信頼済みサイト]** ゾーンを選択して、**[サイト]** を選択します。
3. エラー メッセージに表示される URL を入力して、**[追加]** を選択します。
4. Try to use the add-in again. If the problem persists, verify the settings for the other security zones and ensure that the add-in domain is in the same zone as the URL that is displayed in the address bar of the Office application.

This issue occurs when the Dialog API is used in pop-up mode. To prevent this issue from occurring, use the [displayInFrame](/javascript/api/office/office.ui) flag. This requires that your page support display within an iframe. The following example shows how to use the flag.

```js
Office.context.ui.displayDialogAsync(startAddress, {displayInIFrame:true}, callback);
```

## <a name="changes-to-add-in-commands-including-ribbon-buttons-and-menu-items-do-not-take-effect"></a>リボン ボタンとメニュー項目が含まれているアドイン コマンドへの変更が反映されない

リボン ボタンのアイコンのファイル名やメニュー アイテムのテキストなど、マニフェスト ファイルに変更を加えたときに、変更内容が反映されていないと思われる場合は、そのコンピューターで Office のキャッシュをクリアしてみてください。 

#### <a name="for-windows"></a>Windows の場合:
フォルダー `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` の内容を削除する

#### <a name="for-mac"></a>Mac の場合: 

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

#### <a name="for-ios"></a>iOS の場合: 
Call `window.location.reload(true)` from JavaScript in the add-in to force a reload. Alternatively, you can reinstall Office.

## <a name="changes-to-static-files-such-as-javascript-html-and-css-do-not-take-effect"></a>JavaScript、HTML、CSS などの静的ファイルへの変更は有効になりません

ブラウザーがこれらのファイルをキャッシュしている可能性があります。 これを防ぐには、開発時にクライアント側のキャッシュをオフにします。 詳細は、使用しているサーバーの種類によって異なります。 ほとんどの場合、HTTP 応答に特定のヘッダーを追加する必要があります。 次の設定をお勧めします。

- Cache Control: 「プライベート、キャッシュなし、ストアなし」
- Pragma: 「no-cache」
- 有効期限: 「-1」

Node.JS Express サーバーでこれを行う例については、「[この app.js ファイル](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Complete/app.js)について」を参照してください。 ASP.NET プロジェクトの例については、「[この cshtml ファイル](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Views/Shared/_Layout.cshtml)について」を参照してください。

アドインがインターネット インフォメーション サービス (IIS) にホストされている場合は、次を web.config に追加することもできます。

```xml
<system.webServer>
  <staticContent>
    <clientCache cacheControlMode="UseMaxAge" cacheControlMaxAge="0.00:00:00" cacheControlCustom="must-revalidate" />
  </staticContent>
```

これらの手順が最初に動作しない場合は、ブラウザーのキャッシュをクリアする必要がある場合があります。 これは、ブラウザーの UI を使用して行います。 画面の端の UI でエッジ キャッシュをクリアしようとすると、正常にクリアされないことがあります。 その場合は、Windows コマンド プロンプトで次のコマンドを実行します。

```bash
del /s /f /q %LOCALAPPDATA%\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\#!123\INetCache\
```

## <a name="see-also"></a>関連項目

- [Office on the web でアドインをデバッグする](debug-add-ins-in-office-online.md) 
- [iPad または Mac で Office アドインをサイドロードする](sideload-an-office-add-in-on-ipad-and-mac.md)  
- [iPad と Mac で Office アドインをデバッグする](debug-office-add-ins-on-ipad-and-mac.md)  
- [Visual Studio Code の Microsoft Office アドインデバッガーの拡張機能](./debug-with-vs-extension.md)
- [Office アドインのマニフェストを検証する](troubleshoot-manifest.md)
- [ランタイム ログを使用してアドインをデバッグする](runtime-logging.md)
