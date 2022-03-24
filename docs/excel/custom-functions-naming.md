---
title: カスタム関数の名前付けExcel
description: カスタム関数の名前に関する要件Excel、一般的な名前付けの落とし穴を回避します。
ms.date: 07/08/2021
ms.localizationpriority: medium
ms.openlocfilehash: 629ed7000046a2cf543e0ac9e398c349666a67c1
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744514"
---
# <a name="custom-functions-naming-guidelines"></a>カスタム関数の名前付けのガイドライン

カスタム関数は、JSON メタデータ ファイルの `id` and `name` プロパティによって識別されます。

- この関数は `id` 、JavaScript コード内のカスタム関数を一意に識別するために使用されます。
- この関数`name`は、ユーザーに表示される表示名として使用Excel。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

関数は `name` 、ローカライズの目的など `id`、関数とは異なる場合があります。 一般に、関数`name``id`が異なる理由がない場合と同じにしてください。

関数といくつかの共通`name``id`の要件を共有します。

- 関数では、文字 `id` A から Z、数字 0 から 9、アンダースコア、およびピリオドのみを使用できます。

- 関数では、任意 `name` の Unicode アルファベット文字、アンダースコア、およびピリオドを使用できます。

- どちらの関数も `name` 、 `id` 文字で始まる必要があります。最小制限は 3 文字です。

Excelは、組み込みの関数名 (など) に大文字を使用します`SUM`。 カスタム関数の場合は大文字を使用し、 `name` ベスト プラクティス `id` として使用します。

関数は、 `name` 次の関数と同じにすべきではありません。

- A1 ~ XFD1048576 または R1C1 から R1048576C16384 の間の任意のセル。

- すべてのExcel 4.0 マクロ関数 (`RUN``ECHO`など)  これらの関数の完全な一覧については、「[マクロ関数リファレンス」Excelを参照してください](https://d13ot9o61jdzpp.cloudfront.net/files/Excel%204.0%20Macro%20Functions%20Reference.pdf)。

## <a name="naming-conflicts"></a>名前付けの競合

関数が既 `name` に存在する `name` アドインの関数と同じ場合は、#REF **!** エラーがブックに表示されます。

名前付けの競合を修正するには、アドインの `name` 名前を変更して、もう一度関数を試してください。 また、競合する名前を持つアドインをアンインストールできます。 または、異なる環境でアドインをテストする場合は、別の名前空間を使用して関数 (など) を区別してみてください `NAMESPACE_NAMEOFFUNCTION`。

## <a name="best-practices"></a>ベスト プラクティス

- 同じ名前または類似の名前を持つ複数の関数を作成するのではなく、複数の引数を関数に追加する方法を検討してください。
- 関数名のあいまいな省略形を避ける。 明快さは、明快さよりも重要です。 ではなく、名前 `=INCREASETIME` を選択します `=INC`。
- 関数名は、ZIPCODE ではなく =GETZIPCODE など、関数の動作を示す必要があります。
- 同様のアクションを実行する関数には、同じ動詞を一貫して使用します。 たとえば、and と `=DELETEZIPCODE` `=DELETEADDRESS`、 ではなく 、 を使用 `=DELETEZIPCODE` します `=REMOVEADDRESS`。
- ストリーミング関数に名前を付ける `STREAM` 場合は、関数の説明にメモを追加するか、関数の名前の末尾に追加します。

[!include[manifest guidance](../includes/manifest-guidance.md)]

## <a name="localizing-function-names"></a>関数名のローカライズ

個別の JSON ファイルを使用して、さまざまな言語の関数名をローカライズし、アドインのマニフェスト ファイルの値を上書きできます。 ローカライズされた関数と`id``name`競合する可能性Excel、関数に別の言語で組み込みの関数を与えることは避ける必要があります。

ローカライズの詳細については、「カスタム関数の [ローカライズ」を参照してください。](custom-functions-localize.md)

## <a name="next-steps"></a>次の手順

エラー処理 [のベスト プラクティスについて説明します](custom-functions-errors.md)。

## <a name="see-also"></a>関連項目

* [カスタム関数の JSON メタデータを手動で作成する](custom-functions-json.md)
* [Excel カスタム関数のチュートリアル](../tutorials/excel-tutorial-create-custom-functions.md)
