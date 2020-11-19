---
title: Office アドインの作業ウィンドウ
description: 作業ウィンドウにより、ユーザーはコードを実行してドキュメントや電子メールを修正したり、データ ソースからデータを表示したりするインターフェイス コントロールにアクセスできます。
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: ed3f3b8fdf7cf62b6016fe8b03393de0d56dfb33
ms.sourcegitcommit: ceb8dd66f3fb9c963fce8446c2f6c65ead56fbc1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/18/2020
ms.locfileid: "49132019"
---
# <a name="task-panes-in-office-add-ins"></a>Office アドインの作業ウィンドウ

作業ウィンドウは、通常 Word、PowerPoint、Excel、Outlook 内のウィンドウの右側に表示されるインターフェイスのサーフェスです。作業ウィンドウにより、ユーザーはコードを実行してドキュメントや電子メールを修正したり、データ ソースからデータを表示したりするインターフェイス コントロールにアクセスできます。機能を直接ドキュメントに埋め込む必要がない場合は、作業ウィンドウを使用します。

*図 1. 一般的な作業ウィンドウのレイアウト*

![上部にセクションタブ、左側には会社のロゴと会社名、右下に設定アイコンが表示される一般的な作業ウィンドウレイアウトの図](../images/overview-with-app-task-pane.png)

## <a name="best-practices"></a>ベスト プラクティス

|するべきこと|してはいけないこと|
|:-----|:--------|
|<ul><li>タイトルにアドインの名前を含めます。</li></ul>|<ul><li>タイトルには会社名を追加しません。</li></ul>|
|<ul><li>タイトルには短くわかりやすい名前を使用します。</li></ul>|<ul><li>アドインのタイトルには、「アドイン」、「Word」、「Office」などの文字列を追加しないでください。</li></ul>|
|<ul><li>アドインの上部に CommandBar や Pivot などのナビゲーション要素やコマンド要素を含めます。</li></ul>||
|<ul><li>アドインを Outlook 内で使用する場合を除き、アドインの下部に BrandBar などのブランド化の要素を含めます。</li></ul>||

## <a name="variants"></a>バリアント

次の画像は、1366x768 の解像度で Office アプリのリボンを使用して、さまざまな作業ウィンドウのサイズを示しています。Excel では、数式バーに対応するために、さらに広い領域が必要になります。  

*図 2. Office 2016 デスクトップ作業ウィンドウのサイズ*

![1366x768 の解像度でデスクトップ作業ウィンドウのサイズを表示する図](../images/office-2016-taskpane-sizes.png)

- Excel-320x455 ピクセル
- PowerPoint-320x531 ピクセル
- Word-320x531 ピクセル
- Outlook-348x535 ピクセル

<br/>

*図3Office の作業ウィンドウのサイズ*

![1366x768 の解像度で作業ウィンドウのサイズを表示する図](../images/office-365-taskpane-sizes.png)

- Excel-350x378 ピクセル
- PowerPoint-348x391 ピクセル
- ワード329x445 ピクセル
- Outlook (web 上)-320x570 ピクセル

## <a name="personality-menu"></a>パーソナル メニュー

パーソナル メニューは、アドインの右上付近にあるナビゲーション要素やコマンド要素の妨げになる可能性があります。Windows と Mac でのパーソナル メニューの現在のサイズを次に示します。

Windows の場合、パーソナル メニューは 12x32 ピクセルを測定します (図を参照)。

*図 4. Windows のパーソナル メニュー*

![Windows デスクトップのパーソナルメニューを示す図](../images/personality-menu-win.png)

Mac の場合、パーソナル メニューは 26x26 ピクセルを測定しますが、右から 8 ピクセル内側、上から 6 ピクセルの位置にフロートします。これにより、スペースは 34x32 ピクセルに増加します (図を参照)。

*図 5. Mac のパーソナル メニュー*

![Mac デスクトップのパーソナルメニューを示す図](../images/personality-menu-mac.png)

## <a name="implementation"></a>実装

作業ウィンドウを実装するサンプルについては、GitHub の「[Excel アドインの JS WoodGrove Expense Trends](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends)」を参照してください。

## <a name="see-also"></a>関連項目

- [Office アドインでの Office UI Fabric](office-ui-fabric.md)
- [Office アドインの UX 設計パターン](../design/ux-design-pattern-templates.md)
