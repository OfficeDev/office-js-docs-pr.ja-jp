---
title: Office アドインの作業ウィンドウ
description: 作業ウィンドウにより、ユーザーはコードを実行してドキュメントや電子メールを修正したり、データ ソースからデータを表示したりするインターフェイス コントロールにアクセスできます。
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: cd8d9386fa9f154d611926add12e21f545e36351
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076305"
---
# <a name="task-panes-in-office-add-ins"></a>Office アドインの作業ウィンドウ

作業ウィンドウは、通常 Word、PowerPoint、Excel、Outlook 内のウィンドウの右側に表示されるインターフェイスのサーフェスです。作業ウィンドウにより、ユーザーはコードを実行してドキュメントや電子メールを修正したり、データ ソースからデータを表示したりするインターフェイス コントロールにアクセスできます。機能を直接ドキュメントに埋め込む必要がない場合は、作業ウィンドウを使用します。

*図 1. 一般的な作業ウィンドウのレイアウト*

![上部にセクション タブ、左下に会社のロゴと会社名、右下に設定アイコンを含む一般的な作業ウィンドウ レイアウトを表示する図。](../images/overview-with-app-task-pane.png)

## <a name="best-practices"></a>ベスト プラクティス

|するべきこと|してはいけないこと|
|:-----|:--------|
|<ul><li>タイトルにアドインの名前を含めます。</li></ul>|<ul><li>タイトルには会社名を追加しません。</li></ul>|
|<ul><li>タイトルには短くわかりやすい名前を使用します。</li></ul>|<ul><li>アドインのタイトルには、"アドイン"、"for Word"、"for Office" などの文字列を追加しません。</li></ul>|
|<ul><li>アドインの上部に CommandBar や Pivot などのナビゲーション要素やコマンド要素を含めます。</li></ul>||
|<ul><li>アドインを Outlook 内で使用する場合を除き、アドインの下部に BrandBar などのブランド化の要素を含めます。</li></ul>||

## <a name="variants"></a>バリアント

次の図は、1366x768 解像度のリボンOffice アプリ作業ウィンドウのサイズを示しています。 Excel では、数式バーを収容するための縦のスペースが必要です。  

*図 2. Office 2016 デスクトップ作業ウィンドウのサイズ*

![デスクトップ作業ウィンドウのサイズを 1366x768 解像度で表示する図。](../images/office-2016-taskpane-sizes.png)

- Excel - 320x455 ピクセル
- PowerPoint - 320x531 ピクセル
- Word - 320x531 ピクセル
- Outlook - 348x535 ピクセル

<br/>

*図 3.Office作業ウィンドウのサイズ*

![作業ウィンドウのサイズを 1366x768 解像度で表示する図。](../images/office-365-taskpane-sizes.png)

- Excel - 350x378 ピクセル
- PowerPoint - 348x391 ピクセル
- Word - 329x445 ピクセル
- Outlook (web 上) - 320x570 ピクセル

## <a name="personality-menu"></a>パーソナル メニュー

パーソナル メニューは、アドインの右上付近にあるナビゲーション要素やコマンド要素の妨げになる可能性があります。Windows と Mac でのパーソナル メニューの現在のサイズを次に示します。

Windows の場合、パーソナル メニューは 12x32 ピクセルを測定します (図を参照)。

*図 4. Windows のパーソナル メニュー*

![デスクトップ上のパーソナリティ メニュー Windows図。](../images/personality-menu-win.png)

Mac の場合、パーソナル メニューは 26x26 ピクセルを測定しますが、右から 8 ピクセル内側、上から 6 ピクセルの位置にフロートします。これにより、スペースは 34x32 ピクセルに増加します (図を参照)。

*図 5. Mac のパーソナル メニュー*

![Mac デスクトップのパーソナリティ メニューを示す図。](../images/personality-menu-mac.png)

## <a name="implementation"></a>実装

作業ウィンドウを実装するサンプルについては、GitHub の「[Excel アドインの JS WoodGrove Expense Trends](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends)」を参照してください。

## <a name="see-also"></a>関連項目

- [Office アドインの Fabric Core](fabric-core.md)
- [Office アドインの UX 設計パターン](../design/ux-design-pattern-templates.md)
