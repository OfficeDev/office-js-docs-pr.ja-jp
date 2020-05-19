---
ms.date: 05/17/2020
description: Excel カスタム関数の名前の要件について説明し、一般的な名前付けの落とし穴を回避します。
title: Excel のカスタム関数の名前付けガイドライン
localization_priority: Normal
ms.openlocfilehash: 82b847ba5d944efed16aa2567eee2c3d257a6a75
ms.sourcegitcommit: 54e2892c0c26b9ad1e4dba8aba48fea39f853b6c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/18/2020
ms.locfileid: "44275988"
---
# <a name="naming-guidelines"></a>名前付けのガイドライン

カスタム関数は、 `id` `name` JSON メタデータファイルのおよびプロパティによって識別されます。

- この関数 `id` は、JavaScript コードのカスタム関数を一意に識別するために使用されます。
- 関数 `name` は、Excel でユーザーに表示される表示名として使用されます。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

関数は、 `name` ローカライズのためなど、関数とは異なる場合が `id` あります。 通常、関数は `name` 、 `id` それらを区別する理由がない場合は、と同じです。

`name` `id` いくつかの一般的な要件を共有します。

- 関数では `id` 、a ~ Z の文字を使用することはできません。数字 0 ~ 9、アンダースコア、ピリオド。

- 関数では、 `name` Unicode のアルファベット文字、アンダースコア、ピリオドを使用できます。

- どちらの関数も、 `name` `id` 文字で始まる必要があり、最小で3文字の制限があります。

Excel は、組み込み関数名 (など) に大文字を使用 `SUM` します。 カスタム関数の大文字を使用し `name` 、 `id` ベストプラクティスとして使用します。

関数は `name` 、次のようなものである必要があります。

- A1 から XFD1048576 のセル、または R1C1 から R1048576C16384 までのセル。

- 任意の Excel 4.0 マクロ関数 ( `RUN` 、など `ECHO` )。  これらの関数の完全な一覧については、「 [Excel マクロ関数リファレンスドキュメント](https://d13ot9o61jdzpp.cloudfront.net/files/Excel%204.0%20Macro%20Functions%20Reference.pdf)」を参照してください。

## <a name="naming-conflicts"></a>名前付けの競合

関数 `name` が `name` 既に存在するアドインの関数と同じ場合は、 **#REF!** エラーがブックに表示されます。

名前付けの競合を修正するには、アドインでを変更して、関数を再度実行し `name` ます。 競合する名前を使用してアドインをアンインストールすることもできます。 または、別の環境でアドインをテストしている場合は、別の名前空間を使用して、関数を区別します (など `NAMESPACE_NAMEOFFUNCTION` )。

## <a name="best-practices"></a>ベスト プラクティス

- 同じまたは似た名前を持つ複数の関数を作成するのではなく、複数の引数を関数に追加することを検討してください。
- 関数名にあいまいな略語を含めないでください。 わかりやすくすることが重要です。 ではなく、という名前を選択し `=INCREASETIME` `=INC` ます。
- 関数名は、関数のアクション (ZIPCODE ではなく = GETZIPCODE など) を示す必要があります。
- 類似のアクションを実行する関数に対して同じ動詞を一貫して使用します。 たとえば、とで `=DELETEZIPCODE` はなくを使用し `=DELETEADDRESS` `=DELETEZIPCODE` `=REMOVEADDRESS` ます。
- ストリーミング関数の名前を指定するときは、その効果にメモを追加するか、関数の `STREAM` 名前の末尾に追加することを検討してください。

[!include[manifest guidance](../includes/manifest-guidance.md)]

## <a name="localizing-function-names"></a>関数名のローカライズ

個別の JSON ファイルを使用し、アドインのマニフェストファイルで値をオーバーライドすることにより、異なる言語の関数名をローカライズできます。 `id`ローカライズされた関数と競合する可能性があるため、関数に、または `name` 別の言語の組み込みの Excel 関数を付与しないでください。

ローカライズの詳細については、「[カスタム関数をローカライズ](custom-functions-localize.md)する」を参照してください。

## <a name="next-steps"></a>次の手順
[エラー処理のベストプラクティス](custom-functions-errors.md)について説明します。

## <a name="see-also"></a>関連項目

* [カスタム関数のメタデータ](custom-functions-json.md)
* [Excel カスタム関数のチュートリアル](../tutorials/excel-tutorial-create-custom-functions.md)
