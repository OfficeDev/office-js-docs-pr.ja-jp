---
ms.date: 06/18/2019
description: Excel カスタム関数の名前の要件について説明し、一般的な名前付けの落とし穴を回避します。
title: Excel のカスタム関数の名前付けガイドライン
localization_priority: Normal
ms.openlocfilehash: 18858c45224b833c099c021920ff3d568ad2b3c3
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/21/2019
ms.locfileid: "35128312"
---
# <a name="naming-guidelines"></a>名前付けのガイドライン

カスタム関数は、JSON メタデータファイルの**id**および**name**プロパティによって識別されます。

- この関数`id`は、JavaScript コードのカスタム関数を一意に識別するために使用されます。
- 関数`name`は、Excel でユーザーに表示される表示名として使用されます。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

関数`name`は、ローカライズのためなど`id`、関数とは異なる場合があります。 一般的に、関数の`name`違いがない場合は、 `id`関数はと同じにしておく必要があります。

いくつかの`name`一般的`id`な要件を共有します。

- 関数では`id` 、a ~ Z の文字を使用することはできません。数字 0 ~ 9、アンダースコア、ピリオド。

- 関数では`name` 、Unicode のアルファベット文字、アンダースコア、ピリオドを使用できます。

- どちらの`name`関数`id`も、文字で始まる必要があり、最小で3文字の制限があります。

Excel は、組み込み関数名 (など`SUM`) に大文字を使用します。 そのため、カスタム関数`name`に大文字を使用し、 `id`ベストプラクティスとして使用することを検討してください。

関数`name`には、次のような名前を付けることはできません。

- A1 から XFD1048576 のセル、または R1C1 から R1048576C16384 までのセル。

- 任意の Excel 4.0 マクロ関数 ( `RUN`、 `ECHO`など)。  これらの関数の完全な一覧については、「 [Excel マクロ関数リファレンスドキュメント](https://d13ot9o61jdzpp.cloudfront.net/files/Excel%204.0%20Macro%20Functions%20Reference.pdf)」を参照してください。

## <a name="naming-conflicts"></a>名前付けの競合

関数`name`が既に存在するアドインの関数`name`と同じ場合は、 **#REF!** エラーがブックに表示されます。

名前付けの競合を修正するに`name`は、アドインでを変更して、関数を再度実行します。 競合する名前を使用してアドインをアンインストールすることもできます。 または、別の環境でアドインをテストしている場合は、別の名前空間を使用して、関数`NAMESPACE_NAMEOFFUNCTION`を区別します (など)。

## <a name="best-practices"></a>ベスト プラクティス

- 同じまたは似た名前を持つ複数の関数を作成するのではなく、複数の引数を関数に追加することを検討してください。
- 関数名は、ではなく、関数のアクションを`=GETZIPCODE`示して`ZIPCODE`いなければなりません。
- 関数名にあいまいな略語を含めないでください。 わかりやすくすることが重要です。 ではなく、 `=INCREASETIME`という`=INC`名前を選択します。
- 類似のアクションを実行する関数に対して同じ動詞を一貫して使用します。 たとえば、とで`=DELETEZIPCODE`は`=DELETEADDRESS`なく`=DELETEZIPCODE`を使用し`=REMOVEADDRESS`ます。
- ストリーミング関数の名前を指定するときは、その効果にメモを追加するか、関数の`STREAM`名前の末尾に追加することを検討してください。

## <a name="localizing-function-names"></a>関数名のローカライズ

個別の JSON ファイルを使用し、アドインのマニフェストファイルで値をオーバーライドすることにより、異なる言語の関数名をローカライズできます。 これはローカライズされた関数と競合する`id`可能性`name`があるため、ベストプラクティスとして、関数または組み込みの Excel 関数を別の言語で提供しないようにします。

ローカライズの詳細については、「[カスタム関数をローカライズ](custom-functions-localize.md)する」を参照してください。

## <a name="next-steps"></a>次のステップ
[エラー処理のベストプラクティス](custom-functions-errors.md)について説明します。

## <a name="see-also"></a>関連項目

* [カスタム関数のメタデータ](custom-functions-json.md)
* [カスタム関数のベスト プラクティス](custom-functions-best-practices.md)
* [チュートリアル: Excel でカスタム関数を作成します。](../tutorials/excel-tutorial-create-custom-functions.md)
* [Excel カスタム関数のランタイム](custom-functions-runtime.md)
