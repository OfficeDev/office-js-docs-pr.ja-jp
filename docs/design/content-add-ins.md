---
title: コンテンツ Office アドイン
description: ''
ms.date: 12/04/2017
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

## <a name="see-also"></a>関連項目

- [Office アドインでの Office UI Fabric](office-ui-fabric.md) 
- [Office アドインの UX 設計パターン](ux-design-patterns.md)
