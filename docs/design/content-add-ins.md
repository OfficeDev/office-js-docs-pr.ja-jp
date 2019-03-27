---
title: コンテンツ Office アドイン
description: コンテンツ アドインは、Excel または PowerPoint ドキュメントに直接埋め込むことができるサーフェイスです。これでは、ユーザーはコードを実行してドキュメントを修正したり、データ ソースからデータを表示したりするインターフェイス コントロールにアクセスできます。
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: 46268f963545c3f5b7f45b9b590dc772ba37292f
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870542"
---
# <a name="content-office-add-ins"></a>コンテンツ Office アドイン

コンテンツ アドインは、Excel または PowerPoint ドキュメントに直接埋め込むことができるサーフェイスです。 コンテンツ アドインにより、ユーザーはコードを実行してドキュメントを修正したり、データ ソースからデータを表示したりするインターフェイス コントロールにアクセスできます。 機能を直接ドキュメントに埋め込む場合は、コンテンツ アドインを使用します。  

*図 1. コンテンツ アドインの一般的なレイアウト*

![コンテンツ アドインの一般的なレイアウトを表示する画像の例](../images/overview-with-app-content.png)

## <a name="best-practices"></a>ベスト プラクティス

- アドインの上部に CommandBar や Pivot などのナビゲーション要素やコマンド要素を含めます。
- アドインの下部に BrandBar などのブランド化の要素を含めます (Excel、および PowerPoint アドインにのみ適用)。

## <a name="variants"></a>バリエーション

Office デスクトップと Office 365 の Excel、PowerPoint のコンテンツ アドインのサイズはユーザーが指定します。

## <a name="personality-menu"></a>パーソナル メニュー

パーソナル メニューは、アドインの右上付近にあるナビゲーション要素やコマンド要素の妨げになる可能性があります。Windows と Mac でのパーソナル メニューの現在のサイズを次に示します。

Windows の場合、パーソナル メニューは 12x32 ピクセルを測定します (図を参照)。

*図 2. Windows のパーソナル メニュー* 

![Windows デスクトップのパーソナル メニューを示す図](../images/personality-menu-win.png)


Mac の場合、パーソナル メニューは 26x26 ピクセルを測定しますが、右から 8 ピクセル内側、上から 6 ピクセルの位置にフロートします。これにより、占有スペースは 34x32 ピクセルに増加します (図を参照)。

*図 3. Mac のパーソナル メニュー*

![Mac デスクトップのパーソナル メニューを示す図](../images/personality-menu-mac.png)

## <a name="implementation"></a>実装

コンテンツ アドインの実装サンプルについては、GitHub の「[Excel コンテンツ アドイン Humongous Insurance](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance)」を参照してください。

## <a name="support-considerations"></a>サポートに関する考慮事項

- 使用している Office アドインが[特定の Office ホスト プラットフォーム](/office/dev/add-ins/overview/office-add-in-availability)で動作するかどうかを確認します。 
- コンテンツ アドインによっては、Excel または PowerPoint の読み取りと書き込みのためにユーザーがアドインを「信頼」する必要があります。 アドインのマニフェストには、ユーザーに必要とされる[アクセス許可のレベル](/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins)を宣言することができます。  
- コンテンツ アドインは Office 2013 以降のバージョンの Excel および PowerPoint でサポートされています。 Office Web アドインをサポートしていない Office のバージョンでアドインを開くと、アドインはイメージとして表示されます。

## <a name="see-also"></a>関連項目

- [Office アドインのホストとプラットフォームの可用性](/office/dev/add-ins/overview/office-add-in-availability)
- [Office アドインの Office UI Fabric](/office/dev/add-ins/design/office-ui-fabric)
- [Office アドインの UX 設計パターン](/office/dev/add-ins/design/ux-design-pattern-templates)
- [コンテンツ アドインと作業ウィンドウ アドインでの API 使用についてアクセス許可を要求する](/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins)
