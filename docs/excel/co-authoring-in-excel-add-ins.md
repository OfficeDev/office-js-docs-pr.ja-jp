---
title: Excel アドインの共同編集機能
description: オンラインで、Excel、OneDrive、OneDrive for BusinessにSharePointします。
ms.date: 07/08/2021
ms.localizationpriority: medium
---


# <a name="coauthoring-in-excel-add-ins"></a>Excel アドインの共同編集機能  

[共同編集機能](https://support.microsoft.com/office/7152aa8b-b791-414c-a3bb-3024e46fb104)により、複数のユーザーが連携して同じ Excel ブックを同時に編集できるようになります。 ブックのすべての共同編集者は、他の共同編集者がブックを保存するとすぐに、その共同編集者による変更の内容を確認できます。 Excel ブックを共同編集するには、そのブックが OneDrive、OneDrive for Business、SharePoint Online のいずれかに保存されている必要があります。

> [!IMPORTANT]
> このExcel for Microsoft 365、左上隅に AutoSave が表示されます。 [自動保存] をオンにすると、共同編集者はリアルタイムで変更内容を確認できます。 Excel アドインの設計時には、この動作の影響を考慮に入れておいてください。 ユーザーは、Excel ウィンドウの左上隅にあるスイッチで [自動保存] をオフに切り替えることができます。

## <a name="coauthoring-overview"></a>共同編集機能の概要

ブックの内容に変更を加えると、その変更は Excel によってすべての共同編集者間で同期されます。 共同編集者はブックの内容を変更できますが、Excel アドイン内で実行するコードもブックの内容を変更できます。 たとえば、次の JavaScript コードがアドインで実行Office、範囲の値は Contoso に設定されます。

```js
range.values = [['Contoso']];
```

すべての共同編集者間で 'Contoso' が同期されると、同じブックで作業するユーザーまたは実行中のアドインは、新しい範囲の値を認識するようになります。

共同編集機能では、共有ブック内の内容のみが同期されます。 ブックから Excel アドイン内の JavaScript 変数にコピーした値は同期されません。 たとえば、アドインが JavaScript 変数にセルの値 (たとえば 'Contoso') を保存しているときに、そのセルの値を共同編集者が 'Example' に変更すると、同期後に、そのセルの値はすべての共同編集者に対して 'Example' と表示されます。 ただし、JavaScript 変数の値は 'Contoso' に設定されたままです。 さらに、複数の共同編集者が同じアドインを使用しているときに、それぞれの共同編集者が独自に変数をコピーしている場合、その変数のコピーは同期されません。 ブックの内容を使用する変数を使用するときには、その変数を使用する前に、ブック内で更新された値について必ずチェックしてください。

## <a name="use-events-to-manage-the-in-memory-state-of-your-add-in"></a>イベントを使用したアドインのメモリ内の状態の管理

Excel アドインはブックの内容を読み込んで (非表示のワークシートおよび設定オブジェクトからの読み込み)、その内容を変数などのデータ構造に保存できます。 そのようなデータ構造に元の値がコピーされた後でも、共同編集者は元のブックの内容を更新できます。 つまり、データ構造にコピーした値は、ブックの内容と同期されなくなっているということです。 独自のアドインを構築するときには、ブックの内容とデータ構造に保存された値に関して、このような分離があることを必ず考慮に入れてください。

たとえば、カスタム視覚エフェクトを表示するコンテンツ アドインを作成するとします。 カスタム視覚エフェクトの状態は非表示のワークシートに保存することにします。 共同編集者が同じブックを使用する場合、次のシナリオが発生する可能性があります。

- ユーザー A がドキュメント開くと、カスタム視覚エフェクトがブックに表示されます。 カスタム視覚エフェクトは、非表示のワークシートからデータを読み込みます (たとえば、視覚エフェクトの色が青色に設定されているとします)。
- ユーザー B が同じドキュメントを開いて、カスタム視覚エフェクトの変更を始めます。 ユーザー B は、カスタム視覚エフェクトの色を橙色に設定します。 橙色の設定が非表示のワークシートに保存されます。
- ユーザー A の非表示のワークシートが新しい値の橙色で更新されます。
- ユーザー A のカスタム視覚エフェクトは青色のままです。

ユーザー A のカスタム視覚エフェクトが、共同編集者によって非表示のワークシートに加えられた変更に呼応するようにするには、[BindingDataChanged](/javascript/api/office/office.bindingdatachangedeventargs) イベントを使用します。 これにより、共同編集者がブックの内容に加えた変更が、アドインの状態に反映されるようになります。

## <a name="caveats-to-using-events-with-coauthoring"></a>共同編集機能にイベントを使用する際の注意事項

前述したように、シナリオによっては、すべての共同編集者に向けてイベントをトリガーすることで、ユーザー エクスペリエンスが向上します。 ただし、この動作がユーザー エクスペリエンスの低下を招くシナリオも存在することに注意してください。

たとえば、データの入力規則のシナリオでは、一般に、イベントに呼応して UI を表示します。 前のセクションで説明した [BindingDataChanged](/javascript/api/office/office.bindingdatachangedeventargs) イベントは、ローカル ユーザーまたは共同編集者 (リモート) のどちらかがバインディングの範囲内でブックの内容を変更したときに実行されます。 イベントのイベント ハンドラーに `BindingDataChanged` UI が表示される場合、ブックで作業していた変更とは無関係の UI が表示され、ユーザー エクスペリエンスが低下します。 アドインでイベントを使用する場合は、UI の表示を避けるようにしてください。

## <a name="avoid-table-row-coauthoring-conflicts"></a>テーブル行の共同編集の競合を回避する

API の呼び出しによって [`TableRowCollection.add`](/javascript/api/excel/excel.tablerowcollection#excel-excel-tablerowcollection-add-member(1)) 共同編集の競合が発生する可能性がある既知の問題です。 他のユーザーがアドインのブックを編集している間にアドインが実行される可能性がある場合は、その API を使用することをお勧めしません (特に、テーブルまたはテーブルの下の任意の範囲を編集している場合)。 次のガイダンスは、メソッド`TableRowCollection.add`の問題を回避するのに役立ちます (また、ユーザーに更新を求めるExcel黄色のバーをトリガーしないようにします)。

1. の [`Range.values`](/javascript/api/excel/excel.range#excel-excel-range-values-member) 代わりに使用します [`TableRowCollection.add`](/javascript/api/excel/excel.tablerowcollection#excel-excel-tablerowcollection-add-member(1))。 テーブルの直 `Range` 下に値を設定すると、テーブルが自動的に展開されます。 それ以外の場合、API を使用してテーブル行を追加 `Table` すると、共同認証ユーザーのマージ競合が発生します。
1. データ検証が列 [全体に](https://support.microsoft.com/office/29fecbcc-d1b9-42c1-9d76-eff3ce5f7249) 適用されていない限り、テーブルの下のセルにデータ検証ルールが適用される必要はありません。
1. テーブルの下にデータがある場合、アドインは範囲の値を設定する前に処理する必要があります。 空 [`Range.insert`](/javascript/api/excel/excel.range#excel-excel-range-insert-member(1)) の行を挿入するために使用すると、データが移動され、展開テーブルの領域が作成されます。 それ以外の場合は、表の下のセルを上書きするリスクがあります。
1. テーブルに空の行を追加することはできません `Range.values`。 テーブルの直下のセルにデータが存在する場合にのみ、テーブルは自動的に展開されます。 空のテーブル行を追加するには、回避策として一時的なデータまたは非表示の列を使用します。

## <a name="see-also"></a>関連項目

- [Excel (VBA) の共同編集機能について](/office/vba/excel/concepts/about-coauthoring-in-excel)
- [自動保存がアドインとマクロ (VBA) に及ぼす影響](/office/vba/library-reference/concepts/how-autosave-impacts-addins-and-macros)
