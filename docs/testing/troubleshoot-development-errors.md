---
title: Office アドインでの開発エラーのトラブルシューティング
description: アドインの開発エラーをトラブルシューティングするOffice説明します。
ms.date: 09/24/2021
ms.localizationpriority: medium
ms.openlocfilehash: 2a17a9eafd91cd174209b1974eea61715385c0ad
ms.sourcegitcommit: 517786511749c9910ca53e16eb13d0cee6dbfee6
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/29/2021
ms.locfileid: "59990804"
---
# <a name="troubleshoot-development-errors-with-office-add-ins"></a>Office アドインでの開発エラーのトラブルシューティング

アドインの開発中に発生する可能性がある一般的な問題Office次に示します。

> [!TIP]
> 多くの場合、Officeキャッシュをクリアすると、古いコードに関連する問題が修正されます。 これにより、現在のファイル名、メニュー テキスト、その他のコマンド要素を使用して、最新のマニフェストがアップロードされます。 詳細については、「キャッシュをクリア[する」をOfficeしてください](clear-cache.md)。

## <a name="add-in-doesnt-load-in-task-pane-or-other-issues-with-the-add-in-manifest"></a>アドインが作業ウィンドウで読み込まれない、または他のアドイン マニフェストの問題

アドインのマニフェストでの問題をデバッグするには、「[Office アドインのマニフェストを検証する](troubleshoot-manifest.md)」および「[ランタイム ログを使用してアドインをデバッグする](runtime-logging.md)」を参照してください。

## <a name="changes-to-add-in-commands-including-ribbon-buttons-and-menu-items-do-not-take-effect"></a>リボン ボタンとメニュー項目が含まれているアドイン コマンドへの変更が反映されない

キャッシュをクリアすると、アドインのマニフェストの最新バージョンが使用されます。 キャッシュをクリアするにはOfficeキャッシュをクリアする[の手順に従Officeします](clear-cache.md)。 アプリを使用しているOffice on the web、ブラウザーの UI を使用してブラウザーのキャッシュをクリアします。

## <a name="changes-to-static-files-such-as-javascript-html-and-css-do-not-take-effect"></a>JavaScript、HTML、CSS などの静的ファイルへの変更は有効になりません

ブラウザーがこれらのファイルをキャッシュしている可能性があります。 これを防ぐには、開発時にクライアント側のキャッシュをオフにします。 詳細は、使用しているサーバーの種類によって異なります。 ほとんどの場合、HTTP 応答に特定のヘッダーを追加する必要があります。 次のセットをお勧めします。

- Cache Control: 「プライベート、キャッシュなし、ストアなし」
- Pragma: 「no-cache」
- 有効期限: 「-1」

Node.JS Express サーバーでこれを行う例については、「[この app.js ファイル](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/auth/Office-Add-in-NodeJS-SSO/Complete/app.js)について」を参照してください。 ASP.NET プロジェクトの例については、「[この cshtml ファイル](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/auth/Office-Add-in-ASPNET-SSO/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Views/Shared/_Layout.cshtml)について」を参照してください。

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

## <a name="changes-made-to-property-values-dont-happen-and-there-is-no-error-message"></a>プロパティ値に加えた変更は発生し、エラー メッセージはありません

プロパティが読み取り専用である場合は、プロパティのリファレンス ドキュメントを参照してください。 また[、JS の TypeScript 定義Office、](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md)読み取り専用のオブジェクト プロパティを指定します。 読み取り専用プロパティを設定しようとすると、書き込み操作はサイレント モードで失敗し、エラーはスローされます。 次の例では、読み取り専用プロパティを誤って設定 [Chart.id。](/javascript/api/excel/excel.chart#id)「一部 [のプロパティを直接設定できない」も参照してください](../develop/application-specific-api-model.md#some-properties-cannot-be-set-directly)。

```js
// This will do nothing, since `id` is a read-only property.
myChart.id = "5";
```

## <a name="getting-error-this-add-in-is-no-longer-available"></a>エラーの取得: "このアドインは使用できなくなりました"

このエラーの原因の一部を次に示します。 その他の原因が見つかった場合は、ページの下部にあるフィードバック ツールを使って教えて下さい。

- アプリケーションを使用しているVisual Studio、サイドローディングに問題がある可能性があります。 ホストとホストのすべてのインスタンスOffice閉じるVisual Studio。 再起動してVisual Studio F5 キーを再度押してみてください。
- アドインのマニフェストは、展開場所 (集中展開、SharePoint、ネットワーク共有など) から削除されています。
- マニフェスト内の [ID 要素](../reference/manifest/id.md) の値は、展開されたコピーで直接変更されています。 何らかの理由でこの ID を変更する場合は、まず Office ホストからアドインを削除してから、元のマニフェストを変更したマニフェストに置き換える必要があります。 多くの場合、元のトレースOffice削除するには、キャッシュをクリアする必要があります。 オペレーティング システム[のキャッシュをOffice方法](clear-cache.md)については、「キャッシュのクリア」の記事を参照してください。
- アドインのマニフェストには、マニフェストの `resid` [[リソース](../reference/manifest/resources.md)] セクションのどこにも定義されていないか、使用する場所とセクションで定義されている場所のスペルが一致しません。 `resid` `<Resources>`
- マニフェストの `resid` どこかに 32 文字を超える属性があります。 属性 `resid` と、セクション内の対応するリソースの属性は `id` `<Resources>` 、32 文字を超えることはできません。
- アドインにはカスタム アドイン コマンドがありますが、それをサポートしないプラットフォームで実行しようとしている。 詳細については、「アドイン コマンド [の要件セット」を参照してください](../reference/requirement-sets/add-in-commands-requirement-sets.md)。

## <a name="add-in-doesnt-work-on-edge-but-it-works-on-other-browsers"></a>アドインは Edge では機能しませんが、他のブラウザーで動作します

「[トラブルシューティングと問題Microsoft Edgeする」を参照してください](../concepts/browsers-used-by-office-web-add-ins.md#troubleshooting-microsoft-edge-issues)。

## <a name="excel-add-in-throws-errors-but-not-consistently"></a>Excelはエラーをスローしますが、一貫して発生しません

考[えられる原因Excel、アドインのトラブルシューティング](../excel/excel-add-ins-troubleshooting.md)に関するページを参照してください。

## <a name="manifest-schema-validation-errors-in-visual-studio-projects"></a>プロジェクトのマニフェスト スキーマ検証Visual Studioエラー

マニフェスト ファイルを変更する必要がある新しい機能を使用している場合は、マニフェスト ファイルで検証エラー Visual Studio。 たとえば、共有 JavaScript ランタイムを実装する要素を追加すると、 `<Runtimes>` 次の検証エラーが表示される場合があります。

**名前空間 ' ' の要素 'Host' に、名前空間 ' に無効な子要素 http://schemas.microsoft.com/office/taskpaneappversionoverrides 'Runtimes' が含 http://schemas.microsoft.com/office/taskpaneappversionoverrides まれている**

この場合は、使用する XSD ファイルVisual Studio最新バージョンに更新できます。 最新のスキーマ バージョンは [[MS-OWEMXML]: 付録 A: 完全な XML スキーマです](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8)。

### <a name="locate-the-xsd-files"></a>XSD ファイルを見つける

1. Visual Studio でプロジェクトを開きます。
1. ソリューション **エクスプローラーで、** ファイルを開manifest.xmlします。 マニフェストは、通常、ソリューションの下の最初のプロジェクトに含まれます。
1. [プロパティ **の表示**  >  **] ウィンドウ**(F4) を選択します。
1. [プロパティ **] ウィンドウで**、省略記号 (...) を選択して XML スキーマ エディター **を開** きます。 ここでは、プロジェクトで使用しているすべてのスキーマ ファイルの正確なフォルダーの場所を確認できます。

### <a name="update-the-xsd-files"></a>XSD ファイルを更新する

1. 更新する XSD ファイルをテキスト エディターで開きます。 検証エラーのスキーマ名は XSD ファイル名に関連付けされます。 たとえば **、TaskPaneAppVersionOverridesV1_0.xsd を開きます**。
1. [[MS-OWEMXML] で](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8)更新されたスキーマを検索します。付録 A: 完全な XML スキーマです。 たとえば、TaskPaneAppVersionOverridesV1_0 [は taskpaneappversionoverrides スキーマ です](/openspecs/office_file_formats/ms-owemxml/82e93ec5-de22-42a8-86e3-353c8336aa40)。
1. テキストをテキスト エディターにコピーします。
1. 更新された XSD ファイルを保存します。
1. 新Visual Studio XSD ファイルの変更を取得するには、次のコマンドを再起動します。

古い追加のスキーマに対して、前のプロセスを繰り返します。

## <a name="see-also"></a>関連項目

- [Office on the web でアドインをデバッグする](debug-add-ins-in-office-online.md)
- [iPad または Mac で Office アドインをサイドロードする](sideload-an-office-add-in-on-ipad-and-mac.md)  
- [Mac で Office アドインをデバッグする](debug-office-add-ins-on-ipad-and-mac.md)  
- [Visual Studio Code 用 Microsoft Office アドイン デバッガー拡張機能](debug-with-vs-extension.md)
- [Office アドインのマニフェストを検証する](troubleshoot-manifest.md)
- [ランタイム ログを使用してアドインをデバッグする](runtime-logging.md)
- [Office アドインでのユーザー エラーのトラブルシューティング](testing-and-troubleshooting.md)
- [Microsoft Q&A (office-js-dev)](/answers/topics/office-js-dev.html)
