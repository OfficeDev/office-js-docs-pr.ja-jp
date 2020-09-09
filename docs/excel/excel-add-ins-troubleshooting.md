---
title: Excel アドインのトラブルシューティング
description: Excel アドインの開発エラーをトラブルシューティングする方法について説明します。
ms.date: 09/08/2020
localization_priority: Normal
ms.openlocfilehash: 1bdd96772d3a221ca3a02e3d5dfcfa16561dd5f1
ms.sourcegitcommit: c6308cf245ac1bc66a876eaa0a7bb4a2492991ac
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/08/2020
ms.locfileid: "47409404"
---
# <a name="troubleshooting-excel-add-ins"></a>Excel アドインのトラブルシューティング

この記事では、Excel に固有の問題のトラブルシューティングについて説明します。 ページの下部にあるフィードバックツールを使用して、記事に追加できるその他の問題を提案してください。

## <a name="api-limitations-when-the-active-workbook-switches"></a>アクティブなブックの切り替え時の API の制限

Excel 用のアドインは、一度に1つのブックを操作することを目的としています。 アドインを実行しているブックとは別のブックがフォーカスを取得すると、エラーが発生することがあります。 これは、フォーカスが変更されたときに、特定のメソッドが呼び出されたときにのみ発生します。

このブックスイッチの影響を受ける Api は次のとおりです。

|Excel JavaScript API | スローされたエラー |
|--|--|
| `Chart.activate` | GeneralException |
| `Range.select` | GeneralException |
| `Table.clearFilters` | GeneralException |
| `Workbook.getActiveCell`  | InvalidSelection|
| `Workbook.getSelectedRange` | InvalidSelection|
| `Workbook.getSelectedRanges`  | InvalidSelection|
| `Worksheet.activate` | GeneralException |
| `Worksheet.delete`  | InvalidSelection|
| `Worksheet.gridlines` | GeneralException |
| `Worksheet.showHeadings` | GeneralException |
| `WorksheetCollection.add` | GeneralException |
| `WorksheetFreezePanes.freezeAt` | GeneralException |
| `WorksheetFreezePanes.freezeColumns` | GeneralException |
| `WorksheetFreezePanes.freezeRows` | GeneralException |
| `WorksheetFreezePanes.getLocationOrNullObject`| GeneralException |
| `WorksheetFreezePanes.unfreeze` | GeneralException |

> [!NOTE]
> これは、Windows または Mac で開いている複数の Excel ブックにのみ適用されます。

## <a name="coauthoring"></a>共同編集

共同編集環境でイベントと共に使用するパターンについては、「 [Excel アドインの共同編集](co-authoring-in-excel-add-ins.md) 」を参照してください。 この記事では、など、特定の Api を使用する場合のマージの競合の可能性についても説明し [`TableRowCollection.add`](/javascript/api/excel/excel.tablerowcollection#add-index--values-) ます。

## <a name="see-also"></a>こちらもご覧ください

- [Office アドインでの開発エラーのトラブルシューティング](../testing/troubleshoot-development-errors.md)
- [Office アドインでのユーザー エラーのトラブルシューティング](../testing/testing-and-troubleshooting.md)
