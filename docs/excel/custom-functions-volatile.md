---
ms.date: 06/17/2019
description: 揮発性およびオフラインのストリーミングカスタム関数を実装する方法について説明します。
title: 関数の揮発性の値
localization_priority: Normal
ms.openlocfilehash: 0edf4071ce366c40300663233f1de318a544169b
ms.sourcegitcommit: 4bf5159a3821f4277c07d89e88808c4c3a25ff81
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/18/2019
ms.locfileid: "35059847"
---
# <a name="volatile-values-in-functions"></a>関数の揮発性の値

Volatile 関数は、セルが計算されるたびに値が変更される関数です。 この値は、関数の引数が変更されていない場合でも変更できます。 これらの関数は、Excel が再計算するたびに再計算を行います。 たとえば、`NOW` 関数を呼び出すセルがあるとします。 `NOW` が呼び出される度に、現在の日付と時刻を自動的に返します。

Excel には、`RAND` や `TODAY` などの組み込み揮発性関数がいくつか含まれています。 Excel のすべての揮発性関数の一覧は、「[揮発性および非揮発性関数](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions)」をご覧ください。

カスタム関数を使用すると、独自の揮発性関数を作成することができます。これは、日付、時刻、乱数、およびモデリングを処理するときに便利です。 たとえば、[モンテカルロモンテカルロシミュレーション](https://en.wikipedia.org/wiki/Monte_Carlo_method)では、最適なソリューションを決定するためにランダムな入力を生成する必要があります。

JSON ファイルの自動生成を選択する場合は、JSDOC comment タグ`@volatile`を使用して揮発性関数を宣言します。 Autogeneration の詳細については、「[カスタム関数の JSON メタデータの作成](custom-functions-json-autogeneration.md)」を参照してください。

揮発性のカスタム関数の例を次に示します。これは6つのサイドダイスの重ね合わせをシミュレートします。

```JS
/**
 * Simulates rolling a 6-sided dice.
 * @customfunction
 * @volatile
 */
function roll6sided(): number {
  return Math.floor(Math.random() * 6) + 1;
}
```

## <a name="next-steps"></a>次のステップ
[カスタム関数に状態を保存](custom-functions-save-state.md)する方法について説明します。

## <a name="see-also"></a>関連項目

* [カスタム関数のパラメータオプション](custom-functions-parameter-options.md)
* [カスタム関数のメタデータ](custom-functions-json.md)
* [Excel でカスタム関数を作成する](custom-functions-overview.md)
