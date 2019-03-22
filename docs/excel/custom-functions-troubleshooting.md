---
ms.date: 03/06/2019
description: Excel のカスタム関数で一般的な問題をトラブルシューティングします。
title: カスタム関数のトラブルシューティング (プレビュー)
localization_priority: Priority
ms.openlocfilehash: ada60fb4184095f194ff425823b04456a7bf0e76
ms.sourcegitcommit: c5daedf017c6dd5ab0c13607589208c3f3627354
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/20/2019
ms.locfileid: "30693761"
---
# <a name="troubleshoot-custom-functions"></a>カスタム関数のトラブルシューティング

カスタム関数を作成してテストするとき、製品でエラーが発生する可能性があります。

問題を解決するには、[ランタイム ログを有効にしてエラーをキャプチャ](#enable-runtime-logging)し、[Excel のネイティブ エラー メッセージ](#check-for-excel-error-messages)を参照します。 また、[SSL 証明書の検証](#verify-ssl-certificates)を正しく行っていない、[promises を未解決のままにしている](#ensure-promises-return)、[関数の関連付け](#associate-your-functions)を忘れる、などの一般的な誤りを確認します。

## <a name="enable-runtime-logging"></a>ランタイム ログを有効にする

Windows 上の Office でアドインをテストする場合は、[ランタイム ログを有効にする](https://docs.microsoft.com/ja-JP/office/dev/add-ins/testing/troubleshoot-manifest#use-runtime-logging-to-debug-your-add-in)必要があります。 ランタイム ログでは、問題解明用に別に作成したログ ファイルに `console.log` ステートメントが配信されます。 ステートメントでは、アドインの XML マニフェスト ファイルに関するエラー、実行時の条件、カスタム関数のインストールなど、さまざまなエラーがカバーされます。  ランタイム ログの詳細については、「[アドインのデバッグにランタイム ログを使用する](https://docs.microsoft.com/ja-JP/office/dev/add-ins/testing/troubleshoot-manifest#use-runtime-logging-to-debug-your-add-in)」をご覧ください。  

### <a name="check-for-excel-error-messages"></a>Excel のエラー メッセージを確認する

Excel には多くの組み込みエラー メッセージがあり、計算エラーが発生するとセルに返されます。 カスタム関数では、`#NULL!`、`#DIV/0!`、`#VALUE!`、`#REF!`、`#NAME?`、`#NUM!`、`#N/A`、`#GETTING_DATA` の各エラー メッセージのみが使用されます。

## <a name="common-issues"></a>一般的な問題

### <a name="my-add-in-wont-load-verify-certifications"></a>アドインが読み込まれない: 証明書を確認する

アドインのインストールが失敗する場合は、アドインをホストしている Web サーバーに対して SSL 証明書が正しく構成されていることを確認します。 通常、SSL 証明書に問題がある場合は、アドインを正しくインストールできなかったことを警告する Excel のエラー メッセージが表示されます。 詳細については、「[自己署名証明書を信頼されたルート証明書として追加する](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md)」をご覧ください。

### <a name="my-functions-wont-load-associate-functions"></a>関数が読み込まれない: 関数を関連付ける

カスタム関数のスクリプト ファイルで、各カスタム関数を、[JSON メタデータ ファイル](custom-functions-json.md)で指定されている ID と関連付ける必要があります。 これを行うには、`CustomFunctions.associate()` メソッドを使用します。 通常、このメソッドの呼び出しは、各関数の後またはスクリプト ファイルの最後に行います。 カスタム関数を関連付けないと、カスタム関数は機能しません。

次の例では、add 関数の後で、関数の名前 `add` と対応する JSON ID `ADD` を関連付けています。

```js
function add(first, second){
  return first + second;
}

CustomFunctions.associate("ADD", add);
```

このプロセスの詳細については、「[関数名を JSON メタデータに関連付ける](https://docs.microsoft.com/ja-JP/office/dev/add-ins/excel/custom-functions-best-practices#associating-function-names-with-json-metadata)」をご覧ください。

### <a name="ensure-promises-return"></a>promise の戻り値を確認する

Excel がカスタム関数の完了を待っているときは、セルに #GETTING_DATA が表示されます。 カスタム関数のコードで promise が返されているのに、promise で結果が返されない場合、Excel は #GETTING_DATA を表示し続けます。 すべての promise でセルに結果が正しく返されていることを、関数で確認します。

## <a name="reporting-feedback"></a>フィードバックの報告

ここに記載されていない問題が発生している場合は、お知らせください。 問題を報告するには 2 つの方法があります。

### <a name="in-excel-on-windows-or-mac"></a>Windows または Mac の Excel で

Windows 用または Mac 用の Excel を使用している場合は、Excel から Office の機能拡張チームにフィードバックを直接報告できます。 これを行うには、**[ファイル]、[フィードバック]、[問題点、改善点の報告]** の順に選択します。 問題点や改善点の報告では、発生した問題を理解するために必要なログが提供されます。

### <a name="in-github"></a>GitHub で

ドキュメント ページの下部にある "コンテンツ フィードバック" 機能を使用するか、[カスタム関数リポジトリに直接新しい問題を記入](https://github.com/OfficeDev/Excel-Custom-Functions/issues)して、発生した問題をお気軽に送信してください。

## <a name="see-also"></a>関連項目

* [カスタム関数のメタデータ](custom-functions-json.md)
* [Excel カスタム関数のランタイム](custom-functions-runtime.md)
* [カスタム関数のベスト プラクティス](custom-functions-best-practices.md)
* [カスタム関数の変更ログ](custom-functions-changelog.md)
* [Excel カスタム関数のチュートリアル](../tutorials/excel-tutorial-create-custom-functions.md)
