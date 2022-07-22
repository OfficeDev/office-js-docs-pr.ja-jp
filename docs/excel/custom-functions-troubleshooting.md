---
ms.date: 06/09/2022
description: Excel カスタム関数に関する一般的な問題のトラブルシューティングを行います。
title: カスタム関数のトラブルシューティング
ms.localizationpriority: medium
ms.openlocfilehash: 89d90b6ee94efac0230933313d2c16b5054dda61
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958560"
---
# <a name="troubleshoot-custom-functions"></a>カスタム関数のトラブルシューティング

カスタム関数を作成してテストするとき、製品でエラーが発生する可能性があります。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

問題を解決するには、[ランタイム ログを有効にしてエラーをキャプチャ](#enable-runtime-logging)し、[Excel のネイティブ エラー メッセージ](#check-for-excel-error-messages)を参照します。 また、[promise を未解決のままにしておく](#ensure-promises-return)など、よくある間違いがないか確認します。

## <a name="debugging-custom-functions"></a>カスタム関数のデバッグ

共有ランタイムを使用するカスタム関数アドインをデバッグするには、「共有 [JavaScript ランタイムを使用するように Office アドインを構成する:デバッグ](../develop/configure-your-add-in-to-use-a-shared-runtime.md#debug)」を参照してください。

共有ランタイムを使用しないカスタム関数アドインをデバッグするには、「 [カスタム関数のデバッグ](custom-functions-debugging.md)」を参照してください。

## <a name="enable-runtime-logging"></a>ランタイム ログを有効にする

Windows 上の Office でアドインをテストする場合は、[ランタイム ログを有効にする](../testing/runtime-logging.md)必要があります。 ランタイム ログでは、問題解明用に別に作成したログ ファイルに `console.log` ステートメントが配信されます。 ステートメントでは、アドインの XML マニフェスト ファイルに関するエラー、実行時の条件、カスタム関数のインストールなど、さまざまなエラーがカバーされます。 ランタイム ログの詳細については、「[ランタイム ログを使用してアドインをデバッグする](../testing/runtime-logging.md)」を参照してください。

### <a name="check-for-excel-error-messages"></a>Excel のエラー メッセージを確認する

Excel には多くの組み込みエラー メッセージがあり、計算エラーが発生するとセルに返されます。 カスタム関数では、`#NULL!`、`#DIV/0!`、`#VALUE!`、`#REF!`、`#NAME?`、`#NUM!`、`#N/A`、`#BUSY!` の各エラー メッセージのみが使用されます。

通常、これらのエラーは、あなたがExcelで既によく見たことがあるかもしれないエラーと対応関係があります。 カスタム関数に固有の例外はわずかにあります。以下に記載します。

- `#NAME`エラーは通常、関数の登録に問題があることを意味します。
- `#N/A`エラーは、登録されている間にその機能を実行できなかったということを示す可能性もあります。 この多くは、`CustomFunctions.associate`コマンドが欠落していることが原因です。
- `#VALUE`エラーは通常、関数のスクリプトファイル内のエラーを示します。
- `#REF!`エラーは、関数名がアドイン内に既に存在するの関数名と同じであることを示している可能性があります。

## <a name="clear-the-office-cache"></a>Office のキャッシュをクリアする

カスタム関数に関する情報はOfficeによってキャッシュされます。 開発中、またカスタム関数を使用して繰り返しリロードしている間は、変更が反映されないことがあります。 Officeのキャッシュをクリアすることでこれを修正できます。 詳細については、「[Office のキャッシュをクリアする](../testing/clear-cache.md)」を参照してください。

## <a name="common-problems-and-solutions"></a>一般的な問題と解決策

### <a name="cant-open-add-in-from-localhost-use-a-local-loopback-exemption"></a>localhost からアドインを開けない: ローカル ループバックの除外を使用する

"localhost からこのアドインを開けない" というエラーが表示された場合は、ローカル ループバックの除外を有効にする必要があります。 方法の詳細については、[このMicrosoft サポート記事](/office/troubleshoot/office-suite-issues/cannot-open-add-in-from-localhost)を参照してください。

### <a name="runtime-logging-reports-typeerror-network-request-failed-on-excel-on-windows"></a>Windows 上の Excel でランタイム ログが「TypeError: Network request failed」と報告する

localhost サーバーへの呼び出し中に[ランタイム ログ](custom-functions-troubleshooting.md#enable-runtime-logging)に「TypeError: Network request failed」というエラーが表示された場合は、ローカル ループバック例外を有効にする必要があります。 方法の詳細については、[このMicrosoft サポート記事](/office/troubleshoot/office-suite-issues/cannot-open-add-in-from-localhost)の *オプション 2* を参照してください。

### <a name="ensure-promises-return"></a>promise の戻り値を確認する

Excelがカスタム関数の完了を待っている間、＃BUSY！と表示されます セル内に。 カスタム関数のコードで promise が返されているのに、promise で結果が返されない場合、Excel は `#BUSY!` を表示し続けます。 すべての promise でセルに結果が正しく返されていることを、関数で確認します。

### <a name="error-the-dev-server-is-already-running-on-port-3000"></a>エラー：開発サーバーはすでにポート3000で実行されています。

`npm start`を実行しているときに、開発サーバーが既にポート3000（またはアドインが使用しているポート）で実行されているというエラーが表示されることがあります。 `npm stop`を実行するか、Node.jsウィンドウを閉じることによって、開発サーバーを停止できます。 場合によっては、開発サーバーの実行が停止するまで数分かかることがあります。

### <a name="my-functions-wont-load-associate-functions"></a>関数が読み込まれない: 関数を関連付ける

JSON が登録されておらず、独自の JSON メタデータを作成した場合、`#VALUE!` エラーが表示されるか、アドインを読み込めないという通知が表示されます。 これは通常、各カスタム関数を [JSON メタデータ ファイル](custom-functions-json.md)で指定されている `id` プロパティと関連付ける必要があります。 これは、関数を使用 `CustomFunctions.associate()` して行われます。 通常、この関数呼び出しは、各関数の後、またはスクリプト ファイルの最後に行われます。 カスタム関数を関連付けないと、カスタム関数は機能しません。

次の例では、add 関数の後で、関数の名前 `add` と対応する JSON ID `ADD` を関連付けています。

```js
/**
 * Add two numbers.
 * @customfunction
 * @param {number} first First number.
 * @param {number} second Second number.
 * @returns {number} The sum of the two numbers.
 */
function add(first, second) {
  return first + second;
}

CustomFunctions.associate("ADD", add);
```

このプロセスの詳細については、「 [関数名と JSON メタデータの関連付け](../excel/custom-functions-json.md#associating-function-names-with-json-metadata)」を参照してください。

## <a name="known-issues"></a>既知の問題

既知の問題は、 [Excel Custom Functions GitHub リポジトリ](https://github.com/OfficeDev/Excel-Custom-Functions/issues)で追跡および報告されます。

## <a name="reporting-feedback"></a>フィードバックの報告

ここに記載されていない問題が発生している場合は、お知らせください。 問題を報告するには 2 つの方法があります。

### <a name="in-excel-on-windows-or-mac"></a>Windows または Mac の Excel で

Windows または Mac で Excel を使用している場合は、Excel から Office の機能拡張チームにフィードバックを直接報告できます。 これを行うには、**[ファイル]、[フィードバック]、[問題点、改善点の報告]** の順に選択します。 問題点や改善点の報告では、発生した問題を理解するために必要なログが提供されます。

### <a name="in-github"></a>GitHub で

ドキュメント ページの下部にある "コンテンツ フィードバック" 機能を使用するか、[カスタム関数リポジトリに直接新しい問題を記入](https://github.com/OfficeDev/Excel-Custom-Functions/issues)して、発生した問題をお気軽に送信してください。

## <a name="next-steps"></a>次の手順

「[XLL ユーザー定義関数と互換性のある、カスタム関数を作成する](make-custom-functions-compatible-with-xll-udf.md)」で方法を確認してください。

## <a name="see-also"></a>関連項目

- [カスタム関数用の JSON メタデータの自動生成](custom-functions-json-autogeneration.md)
- [Excel でカスタム関数を作成する](custom-functions-overview.md)
- [カスタム関数のデバッグ](custom-functions-debugging.md)