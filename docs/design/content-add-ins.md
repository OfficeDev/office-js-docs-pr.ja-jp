---
title: コンテンツ Office アドイン
description: コンテンツ アドインは、Excel または PowerPoint ドキュメントに直接埋め込むことができるサーフェイスです。これでは、ユーザーはコードを実行してドキュメントを修正したり、データ ソースからデータを表示したりするインターフェイス コントロールにアクセスできます。
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: cf4ea46b4b924683756063bb36c3f2ea2b8c6764
ms.sourcegitcommit: ceb8dd66f3fb9c963fce8446c2f6c65ead56fbc1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/18/2020
ms.locfileid: "49132082"
---
# <a name="content-office-add-ins"></a>コンテンツ Office アドイン

コンテンツ アドインは、Excel または PowerPoint ドキュメントに直接埋め込むことができるサーフェイスです。 コンテンツ アドインにより、ユーザーはコードを実行してドキュメントを修正したり、データ ソースからデータを表示したりするインターフェイス コントロールにアクセスできます。 機能を直接ドキュメントに埋め込む場合は、コンテンツ アドインを使用します。  

*図 1. コンテンツ アドインの一般的なレイアウト*

![Office アプリケーションにおけるコンテンツアドインの一般的なレイアウト](../images/overview-with-app-content.png)

## <a name="best-practices"></a>ベスト プラクティス

- アドインの上部に CommandBar や Pivot などのナビゲーション要素やコマンド要素を含めます。
- アドインの下部に BrandBar などのブランド化の要素を含めます (Excel、および PowerPoint アドインにのみ適用)。

## <a name="variants"></a>バリエーション

Office デスクトップと Microsoft 365 の Excel および PowerPoint のコンテンツアドインのサイズはユーザーが指定します。

## <a name="personality-menu"></a>パーソナル メニュー

パーソナル メニューは、アドインの右上付近にあるナビゲーション要素やコマンド要素の妨げになる可能性があります。Windows と Mac でのパーソナル メニューの現在のサイズを次に示します。

Windows の場合、パーソナル メニューは 12x32 ピクセルを測定します (図を参照)。

*図 2. Windows のパーソナル メニュー*

![Windows デスクトップ上の12x32 のパーソナルメニュー](../images/personality-menu-win.png)

Mac の場合、パーソナル メニューは 26x26 ピクセルを測定しますが、右から 8 ピクセル内側、上から 6 ピクセルの位置にフロートします。これにより、占有スペースは 34x32 ピクセルに増加します (図を参照)。

*図 3. Mac のパーソナル メニュー*

![Mac デスクトップの34x32 ピクセルのパーソナリティメニュー](../images/personality-menu-mac.png)

## <a name="implementation"></a>実装

コンテンツ アドインの実装サンプルについては、GitHub の「[Excel コンテンツ アドイン Humongous Insurance](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance)」を参照してください。

## <a name="support-considerations"></a>サポートに関する考慮事項

- Office アドインが [特定の office アプリケーションまたはプラットフォーム](../overview/office-add-in-availability.md)で機能するかどうかを確認します。
- コンテンツ アドインによっては、Excel または PowerPoint の読み取りと書き込みのためにユーザーがアドインを「信頼」する必要があります。 アドインのマニフェストには、ユーザーに必要とされる[アクセス許可のレベル](../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)を宣言することができます。  
- コンテンツ アドインは Office 2013 以降のバージョンの Excel および PowerPoint でサポートされています。 Office Web アドインをサポートしていない Office のバージョンでアドインを開くと、アドインはイメージとして表示されます。

## <a name="see-also"></a>関連項目

- [Office アドインの Office クライアント アプリケーションとプラットフォームの可用性](../overview/office-add-in-availability.md)
- [Office アドインの Office UI Fabric](../design/office-ui-fabric.md)
- [Office アドインの UX 設計パターン](../design/ux-design-pattern-templates.md)
- [アドインでの API 使用についてアクセス許可を要求する](../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)
