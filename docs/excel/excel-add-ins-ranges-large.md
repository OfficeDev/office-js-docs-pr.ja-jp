---
title: JavaScript API を使用した大きな範囲の読み取りExcel書き込み
description: JavaScript API を使用して大きな範囲を読み取りまたは書きExcel説明します。
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: f9ef9a36aab3b21bbcc3e44c02edbbead209682a75d72393eb77a4aa98925a1a
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2021
ms.locfileid: "57084045"
---
# <a name="read-or-write-to-a-large-range-using-the-excel-javascript-api"></a>JavaScript API を使用した大きな範囲の読み取りExcel書き込み

この記事では、JavaScript API を使用して大きな範囲への読み取りおよび書き込みを処理Excel説明します。

## <a name="run-separate-read-or-write-operations-for-large-ranges"></a>大きな範囲に対して個別の読み取り操作または書き込み操作を実行する

範囲に多数のセル、値、数値形式、または数式が含まれている場合、その範囲で API 操作を実行できない場合があります。 API は常に範囲に要求された操作 (特定のデータを取得または書き込む) を実行しようとしますが、広い範囲に対する読み取りや書き込みの操作は、過剰なリソース使用によるエラーになる場合があります。 このようなエラーを避けるため、広い範囲に対して読み取りや書き取り操作を 1 回で実行するのではなく、その範囲の小さいサブセットに対して個別に読み取りまたは書き込み操作を実行することをお勧めします。

システムの制限の詳細については、「リソースの制限とパフォーマンスの最適化」の「Excel アドイン」セクションを参照Office[してください](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins)。

### <a name="conditional-formatting-of-ranges"></a>範囲の条件付き書式

範囲には、条件に基づいて個々のセルに適用する書式設定を含めることができます。 この詳細については、「[Excel の範囲に条件付き書式を適用する](excel-add-ins-conditional-formatting.md)」を参照してください。

## <a name="see-also"></a>関連項目

- [Office アドインの Excel JavaScript オブジェクト モデル](excel-add-ins-core-concepts.md)
- [JavaScript API を使用してセルExcelする](excel-add-ins-cells.md)
- [JavaScript API を使用して、非バウンド範囲に対する読み取りExcel書き込み](excel-add-ins-ranges-unbounded.md)
- [Excel アドインで複数の範囲を同時に操作する](excel-add-ins-multiple-ranges.md)
