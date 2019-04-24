---
ms.date: 02/08/2019
description: Excel カスタム関数の名前の要件について説明し、一般的な名前付けの落とし穴を回避します。
title: Excel でのカスタム関数の名前付けのガイドライン (プレビュー)
localization_priority: Normal
ms.openlocfilehash: 954753c35d2df59093661e3b8e92adfa1302e595
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32449282"
---
# <a name="naming-guidelines"></a>名前付けのガイドライン

カスタム関数は、JSON メタデータファイルの**id**および**name**プロパティによって識別されます。 関数 id は、JavaScript コードのカスタム関数を一意に識別するために使用されます。 関数名は、Excel でユーザーに表示される表示名として使用されます。 関数の名前は、ローカライズのためなど、関数の ID とは異なる場合があります。 しかし、一般的には、それが異なるという説得力のある理由がない場合は、ID と同じままにしておく必要があります。

関数名と関数 id は、いくつかの一般的な要件を共有します。

- 関数 id は、文字 a ~ Z、0 ~ 9、アンダースコア、およびピリオドのみを使用できます。

- 関数名には、Unicode の英字、アンダースコア、ピリオドを使用できます。

- 文字で始まる必要があり、最小で3文字に制限されています。

Excel は、組み込み関数名 (など`SUM`) に大文字を使用します。 そのため、ベストプラクティスとして、カスタム関数名と関数 id に大文字を使用することを検討してください。

関数名には、次のような名前を付けないでください。

- A1 から XFD1048576 のセル、または R1C1 から R1048576C16384 までのセル。

- 任意の Excel 4.0 マクロ関数 ( `RUN`、 `ECHO`など)。  これらの関数の完全な一覧については、[この記事](https://www.microsoft.com/en-us/download/details.aspx?id=1465)を参照してください。

## <a name="naming-conflicts"></a>名前付けの競合

関数名が既に存在するアドインの関数名と同じである場合、 **#REF!** エラーがブックに表示されます。

名前の競合を修正するには、アドイン内の名前を変更して、関数を再度実行します。 競合する名前を使用してアドインをアンインストールすることもできます。 または、別の環境でアドインをテストしている場合は、別の名前空間を使用して、関数を区別します (NAMESPACE_NAMEOFFUNCTION など)。

また、アドイン内で関数を使用する方法についても検討します。 多くの場合、同じまたは似た名前を持つ複数の関数を作成するのではなく、複数の引数を関数に追加することをお勧めします。

## <a name="see-also"></a>関連項目

* [カスタム関数のメタデータ](custom-functions-json.md)
* [カスタム関数のベスト プラクティス](custom-functions-best-practices.md)
* [チュートリアル: Excel でカスタム関数を作成します。](../tutorials/excel-tutorial-create-custom-functions.md)
* [Excel カスタム関数のランタイム](custom-functions-runtime.md)
