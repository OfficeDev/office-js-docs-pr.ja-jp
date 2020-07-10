---
title: Excel アドインの共同編集機能
description: OneDrive、OneDrive for Business、または SharePoint Online に格納されている Excel ブックの coauthor について説明します。
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 4414bf64f05c29328c63d0857a6e498495712ff1
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093477"
---
# <a name="coauthoring-in-excel-add-ins"></a>Excel アドインの共同編集機能  

With [coauthoring](https://support.office.com/article/Collaborate-on-Excel-workbooks-at-the-same-time-with-co-authoring-7152aa8b-b791-414c-a3bb-3024e46fb104), multiple people can work together and edit the same Excel workbook simultaneously. All coauthors of a workbook can see another coauthor's changes as soon as that coauthor saves the workbook. To coauthor an Excel workbook, the workbook must be stored in OneDrive, OneDrive for Business, or SharePoint Online.

> [!IMPORTANT]
> Microsoft 365 の Excel では、左上隅に [自動保存] があることがわかります。 [自動保存] をオンにすると、共同編集者はリアルタイムで変更内容を確認できます。 Excel アドインの設計時には、この動作の影響を考慮に入れておいてください。 ユーザーは、Excel ウィンドウの左上隅にあるスイッチで [自動保存] をオフに切り替えることができます。

## <a name="coauthoring-overview"></a>共同編集機能の概要

ブックの内容に変更を加えると、その変更は Excel によってすべての共同編集者間で同期されます。 共同編集者はブックの内容を変更できますが、Excel アドイン内で実行するコードもブックの内容を変更できます。 たとえば、次に示す JavaScript のコードを Office アドイン内で実行すると、範囲の値が Contoso になります。

```js
range.values = [['Contoso']];
```
すべての共同編集者間で 'Contoso' が同期されると、同じブックで作業するユーザーまたは実行中のアドインは、新しい範囲の値を認識するようになります。

Coauthoring only synchronizes the content within the shared workbook. Values copied from the workbook to JavaScript variables in an Excel add-in are not synchronized. For example, if your add-in stores the value of a cell (such as 'Contoso') in a JavaScript variable, and then a coauthor changes the value of the cell to 'Example', after synchronization all coauthors see 'Example' in the cell. However, the value of the JavaScript variable is still set to 'Contoso'. Furthermore, when multiple coauthors use the same add-in, each coauthor has their own copy of the variable, which is not synchronized. When you use variables that use workbook content, be sure you check for updated values in the workbook before you use the variable.

## <a name="use-events-to-manage-the-in-memory-state-of-your-add-in"></a>イベントを使用したアドインのメモリ内の状態の管理

Excel add-ins can read workbook content (from hidden worksheets and a setting object), and then store it in data structures such as variables. After the original values are copied into any of these data structures, coauthors can update the original workbook content. This means that the copied values in the data structures are now out of sync with the workbook content. When you build your add-ins, be sure to account for this separation of workbook content and values stored in data structures.

For example, you might build a content add-in that displays custom visualizations. The state of your custom visualizations might be saved in a hidden worksheet. When coauthors use the same workbook, the following scenario can occur:

- User A opens the document and the custom visualizations are shown in the workbook. The custom visualizations read data from a hidden worksheet (for example, the color of the visualizations is set to blue).
- User B opens the same document, and starts modifying the custom visualizations. User B sets the color of the custom visualizations to orange. Orange is saved to the hidden worksheet.
- ユーザー A の非表示のワークシートが新しい値の橙色で更新されます。
- ユーザー A のカスタム視覚エフェクトは青色のままです。

If you want User A's custom visualizations to respond to changes made by coauthors on the hidden worksheet, use the [BindingDataChanged](/javascript/api/office/office.bindingdatachangedeventargs) event. This ensures that changes to workbook content made by coauthors is reflected in the state of your add-in.

## <a name="caveats-to-using-events-with-coauthoring"></a>共同編集機能にイベントを使用する際の注意事項

As described earlier, in some scenarios, triggering events for all coauthors provides an improved user experience. However, be aware that in some scenarios this behavior can produce poor user experiences. 

たとえば、データの入力規則のシナリオでは、一般に、イベントに呼応して UI を表示します。 前のセクションで説明した [BindingDataChanged](/javascript/api/office/office.bindingdatachangedeventargs) イベントは、ローカル ユーザーまたは共同編集者 (リモート) のどちらかがバインディングの範囲内でブックの内容を変更したときに実行されます。 イベントのイベントハンドラーに `BindingDataChanged` ui が表示されている場合、ユーザーには、ブック内で作業していた変更に関連しない ui が表示されるので、ユーザーの操作が低下します。 アドインでイベントを使用する場合は、UI の表示を避けるようにしてください。

## <a name="see-also"></a>関連項目

- [Excel (VBA) の共同編集機能について](/office/vba/excel/concepts/about-coauthoring-in-excel)
- [自動保存がアドインとマクロ (VBA) に及ぼす影響](/office/vba/library-reference/concepts/how-autosave-impacts-addins-and-macros)
