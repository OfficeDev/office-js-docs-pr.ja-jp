---
title: Office アドイン用の Office の UI 要素
description: Office アドインのさまざまな種類の UI 要素の概要について説明します。
ms.date: 12/24/2019
localization_priority: Normal
ms.openlocfilehash: 5b9907924c674ed9db2294621123c394419d0c12
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093764"
---
# <a name="office-ui-elements-for-office-add-ins"></a>Office アドイン用の Office の UI 要素

You can use several types of UI elements to extend the Office UI, including add-in commands and HTML containers. These UI elements look like a natural extension of Office and work across platforms. You can insert your custom web-based code into any of these elements.

次の図は、作成できる Office UI 要素の種類を示しています。

![Office ドキュメントのリボン、タスク ウィンドウ、ダイアログ ボックス上のアドイン コマンドを示す図](../images/add-in-ui-elements.png)

## <a name="add-in-commands"></a>アドイン コマンド

アドイン[コマンド](add-in-commands.md)を使用して、Office アプリリボンにアドインへのエントリポイントを追加します。 コマンドは、JavaScript コードを実行するか、HTML コンテナーを起動することによって、アドインのアクションを開始します。 2 種類のアドイン コマンドを作成できます。

|**コマンドの種類**|**説明**|
|:---------------|:--------------|
|リボンのボタン、メニュー、およびタブ|Use to add custom buttons, menus (dropdowns), or tabs to the default ribbon in Office. Use Buttons and menus to trigger an action in Office. Use tabs to group and organize buttons and menus.|
|コンテキスト メニュー| Use to extend the default context menu. Context menus are displayed when users right-click text in an Office document or a table in Excel.| 

## <a name="html-containers"></a>HTML コンテナー

Use HTML containers to embed HTML-based UI code within Office clients. These web pages can then reference the Office JavaScript API to interact with content in the document. You can create three types of HTML containers.

|**HTML コンテナー**|**説明**|
|:-----------------|:--------------|
|[作業ウィンドウ](task-pane-add-ins.md)|Display custom UI in the right pane of the Office document. Use task panes to allow users to interact with your add-in side-by-side with the Office document.|
|[コンテンツ アドイン](content-add-ins.md)|Display custom UI embedded within Office documents. Use content add-ins to allow users to interact with your add-in directly within the Office document. For example, you might want to show external content such as videos or data visualizations from other sources. |
|[ダイアログ ボックス](dialog-boxes.md)|Display custom UI in a dialog box that overlays the Office document. Use a dialog box for interactions that require focus and more real estate, and do not require a side-by-side interaction with the document.|

## <a name="see-also"></a>関連項目

- [Excel、Word、PowerPoint のアドイン コマンド](add-in-commands.md)
- [作業ウィンドウ](task-pane-add-ins.md)
- [コンテンツ アドイン](content-add-ins.md)
- [ダイアログ ボックス](dialog-boxes.md)
