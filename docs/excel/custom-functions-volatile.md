---
ms.date: 05/03/2019
description: 揮発性およびオフラインのストリーミングカスタム関数を実装する方法について説明します。
title: 関数内の揮発性値
localization_priority: Normal
ms.openlocfilehash: 1ca3edc3de2d9ac5f2171004f89466352c5cfa1e
ms.sourcegitcommit: ff73cc04e5718765fcbe74181505a974db69c3f5
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/06/2019
ms.locfileid: "33627998"
---
# <a name="volatile-values-in-functions"></a>関数内の揮発性値

Volatile 関数は、セルが計算されるたびに値が変更される関数です。 この値は、関数の引数が変更されていない場合でも変更できます。 これらの関数は、Excel が再計算するたびに再計算を行います。 たとえば、`NOW` 関数を呼び出すセルがあるとします。 `NOW` が呼び出される度に、現在の日付と時刻を自動的に返します。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

Excel には、`RAND` や `TODAY` などの組み込み揮発性関数がいくつか含まれています。 Excel のすべての揮発性関数の一覧は、「[揮発性および非揮発性関数](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions)」をご覧ください。

カスタム関数を使用すると、独自の揮発性関数を作成することができます。これは、日付、時刻、乱数、およびモデリングを処理するときに便利です。 たとえば、[モンテカルロモンテカルロシミュレーション](https://en.wikipedia.org/wiki/Monte_Carlo_method
)では、最適なソリューションを決定するためにランダムな入力を生成する必要があります。

JSON ファイルの自動生成を選択する場合は、JSDOC comment タグ`@volatile`を使用して揮発性関数を宣言します。 Autogeneration の詳細については、「[カスタム関数の JSON メタデータの作成](custom-functions-json-autogeneration.md)」を参照してください。

## <a name="next-steps"></a>次の手順
[カスタム関数に状態を保存](custom-functions-save-state.md)する方法について説明します。

## <a name="see-also"></a>関連項目

* [カスタム関数のパラメータオプション](custom-functions-parameter-options.md)
* [カスタム関数のメタデータ](custom-functions-json.md)
* [Excel でカスタム関数を作成する](custom-functions-overview.md)
