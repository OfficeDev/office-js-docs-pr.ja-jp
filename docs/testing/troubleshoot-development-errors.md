---
title: アドインを使用したOfficeエラーのトラブルシューティング
description: アドインの開発エラーをトラブルシューティングするOffice説明します。
ms.date: 01/04/2021
localization_priority: Normal
ms.openlocfilehash: 48216230db4bf90ca53ef10d98786877bd3905c2
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771425"
---
# <a name="troubleshoot-development-errors-with-office-add-ins"></a>アドインを使用したOfficeエラーのトラブルシューティング

## <a name="add-in-doesnt-load-in-task-pane-or-other-issues-with-the-add-in-manifest"></a>アドインが作業ウィンドウで読み込まれない、または他のアドイン マニフェストの問題

アドインのマニフェストでの問題をデバッグするには、「[Office アドインのマニフェストを検証する](troubleshoot-manifest.md)」および「[ランタイム ログを使用してアドインをデバッグする](runtime-logging.md)」を参照してください。

## <a name="changes-to-add-in-commands-including-ribbon-buttons-and-menu-items-do-not-take-effect"></a>リボン ボタンとメニュー項目が含まれているアドイン コマンドへの変更が反映されない

リボン ボタンのアイコンのファイル名やメニュー アイテムのテキストなど、マニフェスト ファイルに変更を加えたときに、変更内容が反映されていないと思われる場合は、そのコンピューターで Office のキャッシュをクリアしてみてください。 

#### <a name="for-windows"></a>Windows の場合:

フォルダーの内容を削除し `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` 、フォルダーの内容が存在する場合は削除 `%userprofile%\AppData\Local\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\#!123\INetCache\` します。

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

## <a name="changes-made-to-property-values-dont-happen-and-there-is-no-error-message"></a>プロパティ値に加えた変更は行われたので、エラー メッセージはありません

プロパティが読み取り専用である場合は、そのプロパティのリファレンス ドキュメントを確認してください。 また、読み取り専用のOffice JS の [TypeScript](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md) 定義も指定します。 読み取り専用プロパティを設定しようとすると、書き込み操作はサイレント モードで失敗し、エラーはスローされます。 次の例では、誤って読み取り専用プロパティの設定を試 [Chart.id。](/javascript/api/excel/excel.chart#id)「一部 [のプロパティを直接設定できない」も参照してください](../develop/application-specific-api-model.md#some-properties-cannot-be-set-directly)。

```js
// This will do nothing, since `id` is a read-only property.
myChart.id = "5";
```

## <a name="getting-error-this-add-in-is-no-longer-available"></a>エラーが表示される: "このアドインは使用できなくなりました"

このエラーの原因の一部を次に示します。 その他の原因が見つかった場合は、ページの下部にあるフィードバック ツールを使用してご連絡ください。

- アプリを使用しているVisual Studio、サイドローディングに問題がある可能性があります。 ホストとホストのすべてのインスタンスOffice閉Visual Studio。 再起動Visual Studio、もう一度 F5 キーを押してみてください。
- アドインのマニフェストは、一元展開、SharePoint カタログ、ネットワーク共有など、展開場所から削除されました。
- マニフェスト内の [ID 要素](../reference/manifest/id.md) の値は、展開されたコピーで直接変更されています。 何らかの理由でこの ID を変更する場合は、まず Office ホストからアドインを削除してから、元のマニフェストを変更されたマニフェストに置き換える必要があります。 多くの場合、元のOfficeトレースを削除するには、キャッシュをクリアする必要があります。 「リボン ボタンや [メニュー項目を](#changes-to-add-in-commands-including-ribbon-buttons-and-menu-items-do-not-take-effect) 含むアドイン コマンドに対する変更は、この記事の前の方では有効ではありません。」セクションを参照してください。
- アドインのマニフェストには、マニフェストの Resources セクションのどこにも定義されていないものがあります。または、アドインが使用される場所とセクションで定義されている場所のスペルが一致しません。 `resid` [](../reference/manifest/resources.md) `resid` `<Resources>`
- マニフェストのどこかに 32 文字を超える `resid` 属性があります。 属性 `resid` と、セクション内の対応するリソースの属性は `id` `<Resources>` 、32 文字を超えることはできません。
- アドインにはカスタム アドイン コマンドがありますが、それをサポートしないプラットフォーム上で実行しようとしている。 詳細については、アドイン コマンド [の要件セットを参照してください](../reference/requirement-sets/add-in-commands-requirement-sets.md)。

## <a name="add-in-doesnt-work-on-edge-but-it-works-on-other-browsers"></a>アドインは Edge では動作しませんが、他のブラウザーで動作します

「Microsoft [Edge の問題のトラブルシューティング」を参照してください](../concepts/browsers-used-by-office-web-add-ins.md#troubleshooting-microsoft-edge-issues)。

## <a name="excel-add-in-throws-errors-but-not-consistently"></a>Excel アドインはエラーをスローしますが、一貫してスローしません

考 [えられる原因については、「Excel アドインの](../excel/excel-add-ins-troubleshooting.md) トラブルシューティング」を参照してください。

## <a name="see-also"></a>関連項目

- [Office on the web でアドインをデバッグする](debug-add-ins-in-office-online.md)
- [iPad または Mac で Office アドインをサイドロードする](sideload-an-office-add-in-on-ipad-and-mac.md)  
- [iPad と Mac で Office アドインをデバッグする](debug-office-add-ins-on-ipad-and-mac.md)  
- [Visual Studio Code 用 Microsoft Office アドイン デバッガー拡張機能](debug-with-vs-extension.md)
- [Office アドインのマニフェストを検証する](troubleshoot-manifest.md)
- [ランタイム ログを使用してアドインをデバッグする](runtime-logging.md)
- [Office アドインでのユーザー エラーのトラブルシューティング](testing-and-troubleshooting.md)
