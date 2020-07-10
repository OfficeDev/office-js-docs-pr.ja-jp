---
title: Office アドインの作業ウィンドウ
description: 作業ウィンドウにより、ユーザーはコードを実行してドキュメントや電子メールを修正したり、データ ソースからデータを表示したりするインターフェイス コントロールにアクセスできます。
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: 39a96f4d5aa63d55f4dcb30d9aeb9e680357aa09
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093757"
---
# <a name="task-panes-in-office-add-ins"></a>Office アドインの作業ウィンドウ
 
Task panes are interface surfaces that typically appear on the right side of the window within Word, PowerPoint, Excel, and Outlook. Task panes give users access to interface controls that run code to modify documents or emails, or display data from a data source. Use task panes when you don't need to embed functionality directly into the document.

*図 1. 一般的な作業ウィンドウのレイアウト*

![一般的な作業ウィンドウのレイアウトを表示するイメージ](../images/overview-with-app-task-pane.png)

## <a name="best-practices"></a>ベスト プラクティス

|**するべきこと**|**してはいけないこと**|
|:-----|:--------|
|<ul><li>タイトルにアドインの名前を含めます。</li></ul>|<ul><li>タイトルには会社名を追加しません。</li></ul>|
|<ul><li>タイトルには短くわかりやすい名前を使用します。</li></ul>|<ul><li>アドインのタイトルには、「アドイン」、「Word」、「Office」などの文字列を追加しないでください。</li></ul>|
|<ul><li>アドインの上部に CommandBar や Pivot などのナビゲーション要素やコマンド要素を含めます。</li></ul>||
|<ul><li>アドインを Outlook 内で使用する場合を除き、アドインの下部に BrandBar などのブランド化の要素を含めます。</li></ul>||


## <a name="variants"></a>バリアント

The following images show the various task pane sizes with the Office app ribbon at a 1366x768 resolution. For Excel, additional vertical space is required to accommodate the formula bar.  

*図 2. Office 2016 デスクトップ作業ウィンドウのサイズ*

![1366x768 のデスクトップ作業ウィンドウのサイズを示す図](../images/office-2016-taskpane-sizes.png)

- Excel - 320x455
- PowerPoint - 320x531
- Word - 320x531
- Outlook - 348x535

<br/>

*図3Office の作業ウィンドウのサイズ*

![1366x768 のデスクトップ作業ウィンドウのサイズを示す図](../images/office-365-taskpane-sizes.png)

- Excel - 350x378
- PowerPoint - 348x391
- Word - 329x445
- Outlook (on the web) - 320x570

## <a name="personality-menu"></a>パーソナル メニュー

Personality menus can obstruct navigational and commanding elements located near the top right of the add-in. The following are the current dimensions of the personality menu on Windows and Mac.

Windows の場合、パーソナル メニューは 12x32 ピクセルを測定します (図を参照)。

*図 4. Windows のパーソナル メニュー*

![Windows デスクトップのパーソナル メニューを示す図](../images/personality-menu-win.png)

Mac の場合、パーソナル メニューは 26x26 ピクセルを測定しますが、右から 8 ピクセル内側、上から 6 ピクセルの位置にフロートします。これにより、スペースは 34x32 ピクセルに増加します (図を参照)。

*図 5. Mac のパーソナル メニュー*

![Mac デスクトップのパーソナル メニューを示す図](../images/personality-menu-mac.png)

## <a name="implementation"></a>実装

作業ウィンドウを実装するサンプルについては、GitHub の「[Excel アドインの JS WoodGrove Expense Trends](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends)」を参照してください。 


## <a name="see-also"></a>関連項目

- [Office アドインでの Office UI Fabric](office-ui-fabric.md) 
- [Office アドインの UX 設計パターン](../design/ux-design-pattern-templates.md)

