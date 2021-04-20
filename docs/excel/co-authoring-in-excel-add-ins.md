---
title: Excel アドインの共同編集機能
description: OneDrive、OneDrive for Business、または SharePoint Online に格納されている Excel ブックの coauthor について説明します。
ms.date: 07/23/2020
localization_priority: Normal
ms.openlocfilehash: 34ef6fbc32c686e49b9720c5249d5046d26a2952
ms.sourcegitcommit: 7d5407d3900d2ad1feae79a4bc038afe50568be0
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/30/2020
ms.locfileid: "46530444"
---
# <a name="coauthoring-in-excel-add-ins"></a>Excel アドインの共同編集機能  

[共同編集機能](https://support.office.com/article/Collaborate-on-Excel-workbooks-at-the-same-time-with-co-authoring-7152aa8b-b791-414c-a3bb-3024e46fb104)により、複数のユーザーが連携して同じ Excel ブックを同時に編集できるようになります。 ブックのすべての共同編集者は、他の共同編集者がブックを保存するとすぐに、その共同編集者による変更の内容を確認できます。 Excel ブックを共同編集するには、そのブックが OneDrive、OneDrive for Business、SharePoint Online のいずれかに保存されている必要があります。

> [!IMPORTANT]
> Microsoft 365 の Excel では、左上隅に [自動保存] があることがわかります。 [自動保存] をオンにすると、共同編集者はリアルタイムで変更内容を確認できます。 Excel アドインの設計時には、この動作の影響を考慮に入れておいてください。 ユーザーは、Excel ウィンドウの左上隅にあるスイッチで [自動保存] をオフに切り替えることができます。

## <a name="coauthoring-overview"></a>共同編集機能の概要

ブックの内容に変更を加えると、その変更は Excel によってすべての共同編集者間で同期されます。 共同編集者はブックの内容を変更できますが、Excel アドイン内で実行するコードもブックの内容を変更できます。 たとえば、次に示す JavaScript のコードを Office アドイン内で実行すると、範囲の値が Contoso になります。

```js
range.values = [['Contoso']];
```

すべての共同編集者間で 'Contoso' が同期されると、同じブックで作業するユーザーまたは実行中のアドインは、新しい範囲の値を認識するようになります。

共同編集機能では、共有ブック内の内容のみが同期されます。 ブックから Excel アドイン内の JavaScript 変数にコピーした値は同期されません。 たとえば、アドインが JavaScript 変数にセルの値 (たとえば 'Contoso') を保存しているときに、そのセルの値を共同編集者が 'Example' に変更すると、同期後に、そのセルの値はすべての共同編集者に対して 'Example' と表示されます。 ただし、JavaScript 変数の値は 'Contoso' に設定されたままです。 さらに、複数の共同編集者が同じアドインを使用しているときに、それぞれの共同編集者が独自に変数をコピーしている場合、その変数のコピーは同期されません。 ブックの内容を使用する変数を使用するときには、その変数を使用する前に、ブック内で更新された値について必ずチェックしてください。

## <a name="use-events-to-manage-the-in-memory-state-of-your-add-in"></a>イベントを使用したアドインのメモリ内の状態の管理

Excel アドインはブックの内容を読み込んで (非表示のワークシートおよび設定オブジェクトからの読み込み)、その内容を変数などのデータ構造に保存できます。 そのようなデータ構造に元の値がコピーされた後でも、共同編集者は元のブックの内容を更新できます。 つまり、データ構造にコピーした値は、ブックの内容と同期されなくなっているということです。 独自のアドインを構築するときには、ブックの内容とデータ構造に保存された値に関して、このような分離があることを必ず考慮に入れてください。

たとえば、カスタム視覚エフェクトを表示するコンテンツ アドインを作成するとします。 カスタム視覚エフェクトの状態は非表示のワークシートに保存することにします。 共同編集者が同じブックを使用するときに、次のシナリオが考えられます。

- ユーザー A がドキュメント開くと、カスタム視覚エフェクトがブックに表示されます。 カスタム視覚エフェクトは、非表示のワークシートからデータを読み込みます (たとえば、視覚エフェクトの色が青色に設定されているとします)。
- ユーザー B が同じドキュメントを開いて、カスタム視覚エフェクトの変更を始めます。 ユーザー B は、カスタム視覚エフェクトの色を橙色に設定します。 橙色の設定が非表示のワークシートに保存されます。
- ユーザー A の非表示のワークシートが新しい値の橙色で更新されます。
- ユーザー A のカスタム視覚エフェクトは青色のままです。

ユーザー A のカスタム視覚エフェクトが、共同編集者によって非表示のワークシートに加えられた変更に呼応するようにするには、[BindingDataChanged](/javascript/api/office/office.bindingdatachangedeventargs) イベントを使用します。 これにより、共同編集者がブックの内容に加えた変更が、アドインの状態に反映されるようになります。

## <a name="caveats-to-using-events-with-coauthoring"></a>共同編集機能にイベントを使用する際の注意事項

前述したように、シナリオによっては、すべての共同編集者に向けてイベントをトリガーすることで、ユーザー エクスペリエンスが向上します。 ただし、この動作がユーザー エクスペリエンスの低下を招くシナリオも存在することに注意してください。

たとえば、データの入力規則のシナリオでは、一般に、イベントに呼応して UI を表示します。 前のセクションで説明した [BindingDataChanged](/javascript/api/office/office.bindingdatachangedeventargs) イベントは、ローカル ユーザーまたは共同編集者 (リモート) のどちらかがバインディングの範囲内でブックの内容を変更したときに実行されます。 イベントのイベントハンドラーに `BindingDataChanged` ui が表示されている場合、ユーザーには、ブック内で作業していた変更に関連しない ui が表示されるので、ユーザーの操作が低下します。 アドインでイベントを使用する場合は、UI の表示を避けるようにしてください。

## <a name="avoiding-table-row-coauthoring-conflicts"></a>表の行の共同編集の競合を回避する

これは、API への呼び出しによって、共同編集の競合が発生する既知の問題です [`TableRowCollection.add`](/javascript/api/excel/excel.tablerowcollection#add-index--values-) 。 他のユーザーがアドインのブックを編集している間にアドインが実行されることを想定している場合 (特に、テーブルまたはテーブルの下の範囲を編集している場合) は、この API の使用はお勧めしません。 次のガイダンスは、メソッドに問題が発生しないようにするために役立ち `TableRowCollection.add` ます。また、ユーザーに更新を求める黄色のバーは表示されないようにします。

1. [`Range.values`](/javascript/api/excel/excel.range#values)の代わりにを使用 [`TableRowCollection.add`](/javascript/api/excel/excel.tablerowcollection#add-index--values-) します。 テーブル `Range` のすぐ下に値を設定すると、テーブルが自動的に展開されます。 それ以外の場合は、api を使用してテーブルの行を追加すると、 `Table` coauth ユーザーのマージ競合が発生します。
1. データの入力規則が列全体に適用されていない限り、テーブルの下のセルには、[データの入力](https://support.microsoft.com/office/apply-data-validation-to-cells-29fecbcc-d1b9-42c1-9d76-eff3ce5f7249)規則が適用されません。
1. テーブルの下にデータがある場合、アドインは範囲値を設定する前に処理する必要があります。 を使用して [`Range.insert`](/javascript/api/excel/excel.range##insert-shift-) 空の行を挿入すると、データが移動され、拡張テーブルのスペースが作成されます。 それ以外の場合は、テーブルの下のセルを上書きするリスクがあります。
1. を使用して、空の行をテーブルに追加することはできません `Range.values` 。 テーブルのすぐ下のセルにデータが存在する場合は、テーブルが自動的に展開されます。 空のテーブル行を追加する回避策として、一時データまたは非表示の列を使用します。

## <a name="see-also"></a>関連項目

- [Excel (VBA) の共同編集機能について](/office/vba/excel/concepts/about-coauthoring-in-excel)
- [自動保存がアドインとマクロ (VBA) に及ぼす影響](/office/vba/library-reference/concepts/how-autosave-impacts-addins-and-macros)
