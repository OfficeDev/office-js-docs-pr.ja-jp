---
title: Office アドインでの開発エラーのトラブルシューティング
description: Office アドインの開発エラーのトラブルシューティング方法について説明します。
ms.date: 07/28/2022
ms.localizationpriority: medium
ms.openlocfilehash: 18236787ad6ffa9139eb95299723c8935d584668
ms.sourcegitcommit: 143ab022c9ff6ba65bf20b34b5b3a5836d36744c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/03/2022
ms.locfileid: "67177666"
---
# <a name="troubleshoot-development-errors-with-office-add-ins"></a>Office アドインでの開発エラーのトラブルシューティング

Office アドインの開発中に発生する可能性がある一般的な問題の一覧を次に示します。

> [!TIP]
> Office キャッシュをクリアすると、古いコードに関連する問題が修正されることがよくあります。 これにより、現在のファイル名、メニュー テキスト、およびその他のコマンド要素を使用して、最新のマニフェストがアップロードされます。 詳細については、「 [Office キャッシュをクリアする](clear-cache.md)」を参照してください。

## <a name="add-in-doesnt-load-in-task-pane-or-other-issues-with-the-add-in-manifest"></a>アドインが作業ウィンドウで読み込まれない、または他のアドイン マニフェストの問題

アドインのマニフェストでの問題をデバッグするには、「[Office アドインのマニフェストを検証する](troubleshoot-manifest.md)」および「[ランタイム ログを使用してアドインをデバッグする](runtime-logging.md)」を参照してください。

## <a name="changes-to-add-in-commands-including-ribbon-buttons-and-menu-items-do-not-take-effect"></a>リボン ボタンとメニュー項目が含まれているアドイン コマンドへの変更が反映されない

キャッシュをクリアすると、アドインのマニフェストの最新バージョンが使用されていることを確認できます。 Office キャッシュをクリアするには、「Office キャッシュを [クリアする」](clear-cache.md)の手順に従います。 Office on the webを使用している場合は、ブラウザーの UI を使用してブラウザーのキャッシュをクリアします。

## <a name="changes-to-static-files-such-as-javascript-html-and-css-do-not-take-effect"></a>JavaScript、HTML、CSS などの静的ファイルへの変更は有効になりません

ブラウザーがこれらのファイルをキャッシュしている可能性があります。 これを防ぐには、開発時にクライアント側のキャッシュをオフにします。 詳細は、使用しているサーバーの種類によって異なります。 ほとんどの場合、HTTP 応答に特定のヘッダーを追加する必要があります。 次のセットをお勧めします。

- Cache Control: 「プライベート、キャッシュなし、ストアなし」
- Pragma: 「no-cache」
- 有効期限: 「-1」

Node.JS Express サーバーでこれを行う例については、「[この app.js ファイル](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-NodeJS-SSO/Complete/app.js)について」を参照してください。 ASP.NET プロジェクトの例については、「[この cshtml ファイル](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-ASPNET-SSO/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Views/Shared/_Layout.cshtml)について」を参照してください。

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

## <a name="changes-made-to-property-values-dont-happen-and-there-is-no-error-message"></a>プロパティ値に加えられた変更は発生せず、エラー メッセージはありません

プロパティが読み取り専用かどうかを確認するには、プロパティのリファレンス ドキュメントを参照してください。 また、Office JS の [TypeScript 定義](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md) では、読み取り専用のオブジェクト プロパティを指定します。 読み取り専用プロパティを設定しようとすると、書き込み操作はサイレント モードで失敗し、エラーはスローされません。 次の例では、読み取り専用プロパティ [Chart.id](/javascript/api/excel/excel.chart#excel-excel-chart-id-member) の設定を誤って試みます。「 [一部のプロパティを直接設定できない](../develop/application-specific-api-model.md#some-properties-cannot-be-set-directly)」も参照してください。

```js
// This will do nothing, since `id` is a read-only property.
myChart.id = "5";
```

## <a name="getting-error-this-add-in-is-no-longer-available"></a>エラーの取得: "このアドインは使用できなくなりました"

このエラーの原因の一部を次に示します。 その他の原因が見つかれば、ページの下部にあるフィードバック ツールを使用してください。

- Visual Studio を使用している場合は、サイドローディングに問題がある可能性があります。 Office ホストと Visual Studio のすべてのインスタンスを閉じます。 Visual Studio を再起動し、もう一度 F5 キーを押してみてください。
- アドインのマニフェストは、一元展開、SharePoint カタログ、ネットワーク共有など、展開場所から削除されました。
- マニフェスト内の [ID](/javascript/api/manifest/id) 要素の値は、デプロイされたコピーで直接変更されています。 何らかの理由でこの ID を変更する場合は、最初に Office ホストからアドインを削除してから、元のマニフェストを変更されたマニフェストに置き換えます。 多くの場合、元のトレースをすべて削除するには、Office キャッシュをクリアする必要があります。 オペレーティング システム [のキャッシュをクリアする](clear-cache.md) 方法については、Office キャッシュのクリアに関する記事を参照してください。
- アドインのマニフェストには、`resid`マニフェストの [[リソース](/javascript/api/manifest/resources)] セクションのどこにも定義されていないものがあります。また、使用する場所とセクション内で定義 **\<Resources\>** されている場所との間の`resid`スペルが一致しません。
- `resid`マニフェストのどこかに、32 文字を超える属性があります。 `resid`属性と`id`、セクション内 **\<Resources\>** の対応するリソースの属性は、32 文字を超えることはできません。
- アドインにはカスタム アドイン コマンドがありますが、それをサポートしていないプラットフォームで実行しようとしています。 詳細については、「 [アドイン コマンドの要件セット](/javascript/api/requirement-sets/common/add-in-commands-requirement-sets)」を参照してください。

## <a name="add-in-doesnt-work-on-edge-but-it-works-on-other-browsers"></a>アドインは Edge では機能しませんが、他のブラウザーで動作します

[Microsoft Edge の問題のトラブルシューティングに関するページを](../concepts/browsers-used-by-office-web-add-ins.md#troubleshoot-microsoft-edge-issues)参照してください。

## <a name="excel-add-in-throws-errors-but-not-consistently"></a>Excel アドインはエラーをスローしますが、一貫したエラーはスローされません

考えられる原因については、「 [Excel アドインのトラブルシューティング](../excel/excel-add-ins-troubleshooting.md) 」を参照してください。

## <a name="manifest-schema-validation-errors-in-visual-studio-projects"></a>Visual Studio プロジェクトのマニフェスト スキーマ検証エラー

マニフェスト ファイルの変更を必要とする新しい機能を使用している場合は、Visual Studio で検証エラーが発生する可能性があります。 たとえば、共有 JavaScript ランタイムを **\<Runtimes\>** 実装する要素を追加すると、次の検証エラーが表示されることがあります。

**名前空間 '' の要素 'Host' に、名前空間 'http://schemas.microsoft.com/office/taskpaneappversionoverrides''の無効な子要素 'Runtimes' がありますhttp://schemas.microsoft.com/office/taskpaneappversionoverrides。**

これが発生した場合は、Visual Studio が使用する XSD ファイルを最新バージョンに更新できます。 最新のスキーマ バージョンは [、[MS-OWEMXML]: 付録 A: 完全な XML スキーマ](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8)にあります。

### <a name="locate-the-xsd-files"></a>XSD ファイルを見つける

1. Visual Studio でプロジェクトを開きます。
1. **ソリューション エクスプローラー** で、manifest.xml ファイルを開きます。 マニフェストは通常、ソリューションの下の最初のプロジェクトにあります。
1. [ **ビュー** > **のプロパティ] ウィンドウ** (F4) を選択します。
1. **[プロパティ] ウィンドウ** で省略記号 (...) を選択して **、XML スキーマ エディターを** 開きます。 ここでは、プロジェクトが使用するすべてのスキーマ ファイルの正確なフォルダーの場所を確認できます。

### <a name="update-the-xsd-files"></a>XSD ファイルを更新する

1. 更新する XSD ファイルをテキスト エディターで開きます。 検証エラーのスキーマ名は、XSD ファイル名に関連付けられます。 たとえば、 **TaskPaneAppVersionOverridesV1_0.xsd を開きます**。
1. [[MS-OWEMXML]: 付録 A: 完全な XML スキーマ](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8)で更新されたスキーマを見つけます。 たとえば、TaskPaneAppVersionOverridesV1_0は [taskpaneappversionoverrides スキーマにあるとします](/openspecs/office_file_formats/ms-owemxml/82e93ec5-de22-42a8-86e3-353c8336aa40)。
1. テキスト をテキスト エディターにコピーします。
1. 更新された XSD ファイルを保存します。
1. Visual Studio を再起動して、新しい XSD ファイルの変更を取得します。

古い追加スキーマについては、前のプロセスを繰り返すことができます。

## <a name="when-working-offline-no-office-apis-work"></a>オフラインで作業している場合、Office API は機能しません

CDN からではなくローカル コピーから Office JavaScript ライブラリを読み込むと、ライブラリが最新でない場合、API の動作が停止することがあります。 しばらくプロジェクトから離れている場合は、ライブラリを再インストールして最新バージョンを取得します。 プロセスは IDE によって異なります。 環境に基づいて、次のいずれかのオプションを選択します。

- **Visual Studio**: [最新の Office JavaScript API ライブラリへの更新に関するページ](../develop/update-your-javascript-api-for-office-and-manifest-schema-version.md)を参照してください。 
- **その他の IDE**: npm パッケージ [@microsoft/office-js](https://www.npmjs.com/package/@microsoft/office-js) および [@types/office-js](https://www.npmjs.com/package/@types/office-js) を参照してください。

## <a name="see-also"></a>関連項目

- [Office on the web でアドインをデバッグする](debug-add-ins-in-office-online.md)
- [Mac で Office アドインをサイドロードする](sideload-an-office-add-in-on-mac.md)  
- [iPad で Office アドインをサイドロードする](sideload-an-office-add-in-on-ipad.md)  
- [Mac で Office アドインをデバッグする](debug-office-add-ins-on-ipad-and-mac.md)  
- [Visual Studio Code 用 Microsoft Office アドイン デバッガー拡張機能](debug-with-vs-extension.md)
- [Office アドインのマニフェストを検証する](troubleshoot-manifest.md)
- [ランタイム ログを使用してアドインをデバッグする](runtime-logging.md)
- [Office アドインでのユーザー エラーのトラブルシューティング](testing-and-troubleshooting.md)
- [Microsoft Q&A (office-js-dev)](/answers/topics/office-js-dev.html)
