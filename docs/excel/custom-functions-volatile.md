---
ms.date: 01/14/2020
description: 揮発性およびオフラインのストリーミング カスタム関数を実装する方法について説明します。
title: 関数の揮発性の値
ms.localizationpriority: medium
ms.openlocfilehash: 401be3e04a7b36a226547175df4311fc653c027a
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744470"
---
# <a name="volatile-values-in-functions"></a>関数の揮発性の値

揮発性関数は、セルが計算されるごとに値が変化する関数です。 関数の引数が変更された場合でも、値は変更できます。 これらの関数は、Excel が再計算するたびに再計算を行います。 たとえば、`NOW` 関数を呼び出すセルがあるとします。 `NOW` が呼び出される度に、現在の日付と時刻を自動的に返します。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

Excel には、`RAND` や `TODAY` などの組み込み揮発性関数がいくつか含まれています。 Excel の揮発性関数の完全なリストは、「[揮発性および非揮発性関数](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions)」を参照してください。

カスタム関数を使用すると、独自の揮発性関数を作成できます。これは、日付、時刻、乱数、およびモデリングを処理するときに役立つ場合があります。 たとえば、 [モンテカルロシミュレーションでは、](https://en.wikipedia.org/wiki/Monte_Carlo_method) 最適なソリューションを決定するためにランダムな入力を生成する必要があります。

JSON ファイルの自動生成を選択する場合は、JSDoc コメント タグを使用して揮発性関数を宣言します `@volatile`。 自動生成の詳細については、「カスタム関数の JSON メタデータの自動生成 [」を参照してください](custom-functions-json-autogeneration.md)。

揮発性のカスタム関数の例を次に示します。これは、6 辺のサイコロの回転をシミュレートします。

![ランダムな値を返すカスタム関数を示す GIF を使用して、6 辺のサイコロのローリングをシミュレートします。](../images/six-sided-die.gif)

```JS
/**
 * Simulates rolling a 6-sided dice.
 * @customfunction
 * @volatile
 */
function roll6sided() {
  return Math.floor(Math.random() * 6) + 1;
}
```

## <a name="next-steps"></a>次の手順
* カスタム関数 [のパラメーター オプションについて説明します](custom-functions-parameter-options.md)。

## <a name="see-also"></a>関連項目

* [カスタム関数の JSON メタデータを手動で作成する](custom-functions-json.md)
* [Excel でカスタム関数を作成する](custom-functions-overview.md)
