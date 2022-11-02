---
title: コンテンツ Office アドイン
description: コンテンツ アドインは、Excel または PowerPoint ドキュメントに直接埋め込むことができるサーフェイスです。これでは、ユーザーはコードを実行してドキュメントを修正したり、データ ソースからデータを表示したりするインターフェイス コントロールにアクセスできます。
ms.date: 05/12/2021
ms.localizationpriority: medium
ms.openlocfilehash: c10893d60f64d875d92aec979a5700630b2cf96c
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810247"
---
# <a name="content-office-add-ins"></a>コンテンツ Office アドイン

コンテンツ アドインは、Excel または PowerPoint ドキュメントに直接埋め込むことができるサーフェイスです。 コンテンツ アドインにより、ユーザーはコードを実行してドキュメントを修正したり、データ ソースからデータを表示したりするインターフェイス コントロールにアクセスできます。 機能を直接ドキュメントに埋め込む場合は、コンテンツ アドインを使用します。  

*図 1. コンテンツ アドインの一般的なレイアウト*

![Office アプリケーションのコンテンツ アドインの一般的なレイアウト。](../images/overview-with-app-content.png)

## <a name="best-practices"></a>ベスト プラクティス

- アドインの上部に CommandBar や Pivot などのナビゲーション要素やコマンド要素を含めます。
- アドインの下部に BrandBar などのブランド化の要素を含めます (Excel、および PowerPoint アドインにのみ適用)。

## <a name="variants"></a>バリエーション

Office デスクトップと Web ブラウザーの Excel および PowerPoint のコンテンツ アドイン サイズは、ユーザーが指定します。

## <a name="personality-menu"></a>パーソナル メニュー

Personality menus can obstruct navigational and commanding elements located near the top right of the add-in. The following are the current dimensions of the personality menu on Windows and Mac.

Windows の場合、パーソナル メニューは 12x32 ピクセルを測定します (図を参照)。

*図 2. Windows のパーソナル メニュー*

![Windows デスクトップの 12 x 32 ピクセルのパーソナリティ メニュー。](../images/personality-menu-win.png)

Mac の場合、パーソナル メニューは 26x26 ピクセルを測定しますが、右から 8 ピクセル内側、上から 6 ピクセルの位置にフロートします。これにより、占有スペースは 34x32 ピクセルに増加します (図を参照)。

*図 3. Mac のパーソナル メニュー*

![Mac デスクトップの 34 x 32 ピクセルのパーソナリティ メニュー。](../images/personality-menu-mac.png)

## <a name="implementation"></a>実装

コンテンツ アドインの実装サンプルについては、GitHub の「[Excel コンテンツ アドイン Humongous Insurance](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance)」を参照してください。

## <a name="support-considerations"></a>サポートに関する考慮事項

- Office アドインが特定の Office [アプリケーションまたはプラットフォーム](/javascript/api/requirement-sets)で動作するかどうかを確認します。
- コンテンツ アドインによっては、Excel または PowerPoint の読み取りと書き込みのためにユーザーがアドインを「信頼」する必要があります。 アドインのマニフェストには、ユーザーに必要とされる[アクセス許可のレベル](../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)を宣言することができます。  
- Content add-ins are supported in Excel and PowerPoint in Office 2013 version and later. If you open an add-in in a version of Office that doesn't support Office web add-ins, the add-in will be displayed as an image.

## <a name="see-also"></a>関連項目

- [Office アドインの Office クライアント アプリケーションとプラットフォームの可用性](/javascript/api/requirement-sets)
- [Office アドインの Fabric Core](fabric-core.md)
- [Office アドインの UX 設計パターン](../design/ux-design-pattern-templates.md)
- [アドインでの API 使用についてアクセス許可を要求する](../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)
