---
ms.date: 05/08/2019
description: Excel のカスタム関数に関する一般的な問題をトラブルシューティングします。
title: カスタム関数のトラブルシューティング
localization_priority: Priority
ms.openlocfilehash: 999b1fb9b89050ab5c6bcf87e1aac9d2fce13702
ms.sourcegitcommit: a99be9c4771c45f3e07e781646e0e649aa47213f
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/11/2019
ms.locfileid: "33952055"
---
# <a name="troubleshoot-custom-functions"></a>カスタム関数のトラブルシューティング

カスタム関数を作成してテストするとき、製品でエラーが発生する可能性があります。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

問題を解決するには、[ランタイム ログを有効にしてエラーをキャプチャ](#enable-runtime-logging)し、[Excel のネイティブ エラー メッセージ](#check-for-excel-error-messages)を参照します。 また、[予約を未解決のままにしたり](#ensure-promises-return)、[機能の関連付け](#my-functions-wont-load-associate-functions)を忘れてしまうといったよくある間違いを確認します。

## <a name="enable-runtime-logging"></a>ランタイム ログを有効にする

Windows 上の Office でアドインをテストする場合は、[ランタイム ログを有効にする](/office/dev/add-ins/testing/troubleshoot-manifest#use-runtime-logging-to-debug-your-add-in)必要があります。 ランタイム ログでは、問題解明用に別に作成したログ ファイルに `console.log` ステートメントが配信されます。 ステートメントでは、アドインの XML マニフェスト ファイルに関するエラー、実行時の条件、カスタム関数のインストールなど、さまざまなエラーがカバーされます。  ランタイム ログの詳細については、「[アドインのデバッグにランタイム ログを使用する](/office/dev/add-ins/testing/troubleshoot-manifest#use-runtime-logging-to-debug-your-add-in)」をご覧ください。  

### <a name="check-for-excel-error-messages"></a>Excel のエラー メッセージを確認する

Excel には多くの組み込みエラー メッセージがあり、計算エラーが発生するとセルに返されます。 カスタム関数では、`#NULL!`、`#DIV/0!`、`#VALUE!`、`#REF!`、`#NAME?`、`#NUM!`、`#N/A`、`#BUSY!` の各エラー メッセージのみが使用されます。

通常、これらのエラーは、あなたがExcelで既によく見たことがあるかもしれないエラーと対応関係があります。 カスタム関数に固有の例外はわずかにあります。以下に記載します。

- `#NAME`エラーは通常、関数の登録に問題があることを意味します。
- `#VALUE`エラーは通常、関数のスクリプトファイル内のエラーを示します。
- `#N/A`エラーは、登録されている間にその機能を実行できなかったということを示す可能性もあります。 この多くは、`CustomFunctions.associate`コマンドが欠落していることが原因です。
- `#REF!`エラーは、関数名がアドイン内に既に存在するの関数名と同じであることを示している可能性があります。

## <a name="clear-the-office-cache"></a>Office のキャッシュをクリアする

カスタム関数に関する情報はOfficeによってキャッシュされます。 開発中、またカスタム関数を使用して繰り返しリロードしている間は、変更が反映されないことがあります。 Officeのキャッシュをクリアすることでこれを修正できます。 詳細については、記事[マニフェストの問題を検証、問題解決する](https://docs.microsoft.com/office/dev/add-ins/testing/troubleshoot-manifest?branch=master#clear-the-office-cache)内「Officeキャッシュをクリアする」の部分を参照してください。

## <a name="common-issues"></a>一般的な問題

### <a name="my-functions-wont-load-associate-functions"></a>関数が読み込まれない: 関数を関連付ける

カスタム関数のスクリプト ファイルで、各カスタム関数を、[JSON メタデータ ファイル](custom-functions-json.md)で指定されている ID と関連付ける必要があります。 これを行うには、`CustomFunctions.associate()` メソッドを使用します。 通常、このメソッドの呼び出しは、各関数の後またはスクリプト ファイルの最後に行います。 カスタム関数を関連付けないと、カスタム関数は機能しません。

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

このプロセスの詳細については、「[関数名を JSON メタデータに関連付ける](/office/dev/add-ins/excel/custom-functions-best-practices#associating-function-names-with-json-metadata)」をご覧ください。

### <a name="cant-open-add-in-from-localhost-use-a-local-loopback-exception"></a>localhostからアドインを開くことができません：ローカルループバック例外を使用してください

"We can't open this add-in from localhost"というエラーが表示された場合は、ローカルループバック例外を有効にする必要があります。 方法の詳細については、[このMicrosoft サポート記事](https://support.microsoft.com/ja-JP/help/4490419/local-loopback-exemption-does-not-work)を参照してください。

### <a name="ensure-promises-return"></a>promise の戻り値を確認する

Excelがカスタム関数の完了を待っている間、＃BUSY！と表示されます セル内に。 カスタム関数のコードで promise が返されているのに、promise で結果が返されない場合、Excel は #BUSY! を表示し続けます。 すべての promise でセルに結果が正しく返されていることを、関数で確認します。

### <a name="error-the-dev-server-is-already-running-on-port-3000"></a>エラー：開発サーバーはすでにポート3000で実行されています。

`npm start`を実行しているときに、開発サーバーが既にポート3000（またはアドインが使用しているポート）で実行されているというエラーが表示されることがあります。 `npm stop`を実行するか、Node.jsウィンドウを閉じることによって、開発サーバーを停止できます。 しかし場合によっては、開発サーバーが実際に実行を停止するのに数分かかることがあります。

## <a name="reporting-feedback"></a>フィードバックの報告

ここに記載されていない問題が発生している場合は、お知らせください。 問題を報告するには 2 つの方法があります。

### <a name="in-excel-on-windows-or-mac"></a>Windows または Mac の Excel で

Windows または Mac で Excel を使用している場合は、Excel から Office の機能拡張チームにフィードバックを直接報告できます。 これを行うには、**[ファイル]、[フィードバック]、[問題点、改善点の報告]** の順に選択します。 問題点や改善点の報告では、発生した問題を理解するために必要なログが提供されます。

### <a name="in-github"></a>GitHub で

ドキュメント ページの下部にある "コンテンツ フィードバック" 機能を使用するか、[カスタム関数リポジトリに直接新しい問題を記入](https://github.com/OfficeDev/Excel-Custom-Functions/issues)して、発生した問題をお気軽に送信してください。

## <a name="next-steps"></a>次の手順
[カスタム関数をデバッグする](custom-functions-debugging.md)手順をご参照ください。

## <a name="see-also"></a>関連項目

* [カスタム関数メタデータ自動生成](custom-functions-json-autogeneration.md)
* [Excel カスタム関数のランタイム](custom-functions-runtime.md)
* [カスタム関数のベスト プラクティス](custom-functions-best-practices.md)
* [カスタム関数をXLLユーザー定義関数と互換性のあるものにします](make-custom-functions-compatible-with-xll-udf.md)。
* [Excel でカスタム関数を作成する](custom-functions-overview.md)
