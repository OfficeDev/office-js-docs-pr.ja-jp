---
title: Office アドインの作業ウィンドウ
description: 作業ウィンドウにより、ユーザーはコードを実行してドキュメントや電子メールを修正したり、データ ソースからデータを表示したりするインターフェイス コントロールにアクセスできます。
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: 048c6a508a9d8106d129469880396a564b26aea4
ms.sourcegitcommit: 6c7c98f085dd20f827e0c388e672993412944851
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/06/2020
ms.locfileid: "42554679"
---
# <a name="task-panes-in-office-add-ins"></a>Office アドインの作業ウィンドウ
 
作業ウィンドウは、通常 Word、PowerPoint、Excel、Outlook 内のウィンドウの右側に表示されるインターフェイスのサーフェスです。作業ウィンドウにより、ユーザーはコードを実行してドキュメントや電子メールを修正したり、データ ソースからデータを表示したりするインターフェイス コントロールにアクセスできます。機能を直接ドキュメントに埋め込む必要がない場合は、作業ウィンドウを使用します。

*図 1. 一般的な作業ウィンドウのレイアウト*

![一般的な作業ウィンドウのレイアウトを表示するイメージ](../images/overview-with-app-task-pane.png)

## <a name="best-practices"></a>ベスト プラクティス

|**するべきこと**|**してはいけないこと**|
|:-----|:--------|
|<ul><li>タイトルにアドインの名前を含めます。</li></ul>|<ul><li>タイトルには会社名を追加しません。</li></ul>|
|<ul><li>タイトルには短くわかりやすい名前を使用します。</li></ul>|<ul><li>アドインのタイトルに “add-in”、“for Word”、“for Office” などの文字列を追加しません。</li></ul>|
|<ul><li>アドインの上部に CommandBar や Pivot などのナビゲーション要素やコマンド要素を含めます。</li></ul>||
|<ul><li>アドインを Outlook 内で使用する場合を除き、アドインの下部に BrandBar などのブランド化の要素を含めます。</li></ul>||


## <a name="variants"></a>バリアント

以下の図は、Office リボンの解像度が 1366x768 のさまざまな作業ウィンドウのサイズを示しています。Excel では、数式バーを収容するための縦のスペースが必要です。  

*図 2. Office 2016 デスクトップ作業ウィンドウのサイズ*

![1366x768 のデスクトップ作業ウィンドウのサイズを示す図](../images/office-2016-taskpane-sizes.png)

- Excel - 320x455
- PowerPoint - 320x531
- Word - 320x531
- Outlook - 348x535

<br/>

*図 3. Office 365 の作業ウィンドウのサイズ*

![1366x768 のデスクトップ作業ウィンドウのサイズを示す図](../images/office-365-taskpane-sizes.png)

- Excel - 350x378
- PowerPoint - 348x391
- Word - 329x445
- Outlook (on the web) - 320x570

## <a name="personality-menu"></a>パーソナル メニュー

パーソナル メニューは、アドインの右上付近にあるナビゲーション要素やコマンド要素の妨げになる可能性があります。Windows と Mac でのパーソナル メニューの現在のサイズを次に示します。

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

