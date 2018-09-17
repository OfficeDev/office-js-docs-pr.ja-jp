---
title: コンテンツ Office アドイン
description: コンテンツのアドインは、Word、Excel に直接埋め込むことができるサーフェスまたはインターフェイスのコントロールにアクセスできるように PowerPoint ドキュメントがドキュメントの変更や、データ ソースからデータを表示するコードを実行します。
ms.date: 12/04/2017
ms.openlocfilehash: 8692b6e8af4504a5eadcba64c9659adaa9122975
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2018
ms.locfileid: "23944084"
---
# <a name="content-office-add-ins"></a>コンテンツ Office アドイン

コンテンツ アドインは、Word、Excel、または PowerPoint ドキュメントに直接埋め込むことができるサーフェイスです。コンテンツ アドインにより、ユーザーはコードを実行してドキュメントを修正したり、データ ソースからデータを表示したりするインターフェイス コントロールにアクセスできます。機能を直接ドキュメントに埋め込む場合は、コンテンツ アドインを使用します。  

*図 1. コンテンツ アドインの一般的なレイアウト*

![コンテンツ アドインの一般的なレイアウトを表示する画像の例](../images/overview-with-app-content.png)

## <a name="best-practices"></a>ベスト プラクティス

- アドインの上部に CommandBar や Pivot などのナビゲーション要素やコマンド要素を含めます。
- アドインの下部に BrandBar などのブランド化の要素を含めます (Word、Excel、および PowerPoint アドインにのみ適用)。

## <a name="variants"></a>バリアント

Office 2016 デスクトップと Office 365 の Word、Excel、PowerPoint のコンテンツ アドインのサイズはユーザーが指定します。

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
- 使用している Office アドインが[特定の Office ホスト プラットフォーム](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-in-availability)で動作するかどうかを確認します。 
- コンテンツ アドインによっては、Excel または PowerPoint の読み取りと書き込みのためにユーザーがアドインを「信頼」する必要があります。 アドインのマニフェストで、ユーザーに必要とされる [アクセス許可のレベル](https://docs.microsoft.com/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins) を宣言することができます。  
- コンテンツ アドインは Office 2013 以降のバージョンの Excel および PowerPoint でサポートされています。 Office Web アドインをサポートしていない Office のバージョンでアドインを開くと、アドインはイメージとして表示されます。

## <a name="see-also"></a>関連項目
- [Office アドインのホストとプラットフォームの可用性](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-in-availability)
- [Office アドインの Office UI Fabric](https://docs.microsoft.com/office/dev/add-ins/design/office-ui-fabric) 
- [Office アドインの UX 設計パターン](https://docs.microsoft.com/office/dev/add-ins/design/ux-design-pattern-templates)
- [コンテンツ アドインと作業ウィンドウ アドインでの API 使用についてアクセス許可を要求する](https://docs.microsoft.com/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins)
