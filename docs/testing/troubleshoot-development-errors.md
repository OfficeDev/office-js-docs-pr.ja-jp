---
title: Office アドインでの開発エラーのトラブルシューティング
description: Office アドインの開発エラーをトラブルシューティングする方法について説明します。
ms.date: 09/08/2020
localization_priority: Normal
ms.openlocfilehash: 5801146165446352ec806f6f832e9976f96467ac
ms.sourcegitcommit: c6308cf245ac1bc66a876eaa0a7bb4a2492991ac
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/08/2020
ms.locfileid: "47409410"
---
# <a name="troubleshoot-development-errors-with-office-add-ins"></a>Office アドインでの開発エラーのトラブルシューティング

## <a name="add-in-doesnt-load-in-task-pane-or-other-issues-with-the-add-in-manifest"></a>アドインが作業ウィンドウで読み込まれない、または他のアドイン マニフェストの問題

アドインのマニフェストでの問題をデバッグするには、「[Office アドインのマニフェストを検証する](troubleshoot-manifest.md)」および「[ランタイム ログを使用してアドインをデバッグする](runtime-logging.md)」を参照してください。

## <a name="changes-to-add-in-commands-including-ribbon-buttons-and-menu-items-do-not-take-effect"></a>リボン ボタンとメニュー項目が含まれているアドイン コマンドへの変更が反映されない

リボン ボタンのアイコンのファイル名やメニュー アイテムのテキストなど、マニフェスト ファイルに変更を加えたときに、変更内容が反映されていないと思われる場合は、そのコンピューターで Office のキャッシュをクリアしてみてください。 

#### <a name="for-windows"></a>Windows の場合:

フォルダーの内容を削除 `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` し、フォルダーの内容を削除し `%userprofile%\AppData\Local\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\#!123\INetCache\` ます (存在する場合)。

#### <a name="for-mac"></a>Mac の場合: 

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

#### <a name="for-ios"></a>iOS の場合: 
アドイン内の JavaScript から `window.location.reload(true)` を呼び出して強制的に再読み込みします。または、Office を再インストールしてください。

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

## <a name="changes-made-to-property-values-dont-happen-and-there-is-no-error-message"></a>プロパティ値に対する変更は行われず、エラーメッセージもありません。

プロパティの参照ドキュメントが読み取り専用かどうかを確認します。 また、Office JS の [TypeScript 定義](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md) は、読み取り専用のオブジェクトプロパティを指定します。 読み取り専用プロパティを設定しようとすると、エラーがスローされずに書き込み操作が失敗します。 次の例では、誤って読み取り専用プロパティ [Chart.id](/javascript/api/excel/excel.chart#id)を設定しようとしています。一部の [プロパティを直接設定することはできません](../develop/application-specific-api-model.md#some-properties-cannot-be-set-directly)。

```js
// This will do nothing, since `id` is a read-only property.
myChart.id = "5";
```

## <a name="add-in-doesnt-work-on-edge-but-it-works-on-other-browsers"></a>アドインはエッジでは動作しませんが、他のブラウザーで動作します。

[Microsoft Edge の問題のトラブルシューティング](../concepts/browsers-used-by-office-web-add-ins.md#troubleshooting-microsoft-edge-issues)を参照してください。

## <a name="excel-add-in-throws-errors-but-not-consistently"></a>Excel アドインはエラーをスローしますが、一貫していません

考えられる原因については、「 [Excel アドインのトラブルシューティング](../excel/excel-add-ins-troubleshooting.md) 」を参照してください。

## <a name="see-also"></a>関連項目

- [Office on the web でアドインをデバッグする](debug-add-ins-in-office-online.md)
- [iPad または Mac で Office アドインをサイドロードする](sideload-an-office-add-in-on-ipad-and-mac.md)  
- [iPad と Mac で Office アドインをデバッグする](debug-office-add-ins-on-ipad-and-mac.md)  
- [Visual Studio Code 用 Microsoft Office アドイン デバッガー拡張機能](debug-with-vs-extension.md)
- [Office アドインのマニフェストを検証する](troubleshoot-manifest.md)
- [ランタイム ログを使用してアドインをデバッグする](runtime-logging.md)
- [Office アドインでのユーザー エラーのトラブルシューティング](testing-and-troubleshooting.md)
