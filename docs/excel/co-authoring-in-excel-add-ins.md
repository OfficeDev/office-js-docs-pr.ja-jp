---
title: Excel アドインの共同編集機能
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 47482ba74ef298add4cdb553d66904381123289d
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2018
ms.locfileid: "23945730"
---
# <a name="coauthoring-in-excel-add-ins"></a>Excel アドインの共同編集機能  

[共同編集機能](https://support.office.com/article/Collaborate-on-Excel-workbooks-at-the-same-time-with-co-authoring-7152aa8b-b791-414c-a3bb-3024e46fb104)により、複数のユーザーが連携して同じ Excel ブックを同時に編集できるようになります。 ブックのすべての共同編集者は、他の共同編集者がブックを保存するとすぐに、その共同編集者による変更の内容を確認できます。 Excel ブックを共同編集するには、そのブックが OneDrive、OneDrive for Business、SharePoint Online のいずれかに保存されている必要があります。

> [!IMPORTANT]
> Office 365のExcel 2016で、左上の角にAutoSaveが表示されます。AutoSaveがオンになっているとき、 共著者がリアルタイムで変化に気づきます。 Excelアドインでのこの動きの影響を考慮してください。ユーザーはExcelウィンドウの左上にあるスイッチでAutoSaveをオフにできます。

共同編集機能は、次のプラットフォームで使用できます。

- Excel Online
- Android用Excel
- iOS 用の Excel
- Windows 10 用の Excel Mobile
- Excel for Windows Desktop for Office 365 のユーザー (Windows デスクトップ ビルド 16.0.8326.2076 以降、2017 年 8 月から Current Channel のユーザーに公開)

## <a name="coauthoring-overview"></a>共同編集機能の概要
 
ワークブックの内容を変更するとき、Excelはすべての共著者を通してそれらの変更を自動的に同期します。共著者はワークブックの内容を変更できますが、Excelアドイン内で実行されているコードもそのようになります。例えば、以下のJavaScriptコードがOffice アドインで実行されているとき、範囲の値はContosoに設定されます:

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

ユーザー A のカスタム視覚エフェクトが、共同編集者によって非表示のワークシートに加えられた変更に呼応するようにするには、[BindingDataChanged](https://docs.microsoft.com/javascript/api/office/office.bindingdatachangedeventargs?view=office-js) イベントを使用します。 これにより、共同編集者がブックの内容に加えた変更が、アドインの状態に反映されるようになります。

## <a name="caveats-to-using-events-with-coauthoring"></a>共同編集機能にイベントを使用する際の注意事項 

前述したように、シナリオによっては、すべての共同編集者に向けてイベントをトリガーすることで、ユーザー エクスペリエンスが向上します。 ただし、この動作がユーザー エクスペリエンスの低下を招くシナリオも存在することに注意してください。 

たとえば、データの入力規則のシナリオでは、一般に、イベントに呼応して UI を表示します。 前のセクションで説明した [BindingDataChanged](https://docs.microsoft.com/javascript/api/office/office.bindingdatachangedeventargs?view=office-js) イベントは、ローカル ユーザーまたは共同編集者 (リモート) のどちらかがバインディングの範囲内でブックの内容を変更したときに実行されます。 **BindingDataChanged** イベントのイベント ハンドラーが UI を表示すると、ユーザーに自分が行った変更とは無関係の UI が表示されるようになるため、ユーザー エクスペリエンスが低下します。 アドインでイベントを使用する場合は、UI の表示を避けるようにしてください。

## <a name="see-also"></a>関連項目 

- [Excel (VBA) の共同編集機能について](https://docs.microsoft.com/office/vba/excel/concepts/about-coauthoring-in-excel) 
- [自動保存がアドインとマクロ (VBA) に及ぼす影響](https://docs.microsoft.com/office/vba/library-reference/concepts/how-autosave-impacts-addins-and-macros) 
