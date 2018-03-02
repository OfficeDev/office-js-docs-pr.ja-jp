---
title: Office アドインでの Office UI Fabric
description: ''
ms.date: 12/04/2017
---


# <a name="office-ui-fabric-in-office-add-ins"></a>Office アドインでの Office UI Fabric 

Office UI Fabric は、Office と Office 365 のユーザー エクスペリエンスをビルドするための JavaScript フロントエンドのフレームワークです。Fabric は、拡張や改訂が可能な視覚効果に焦点を合わせたコンポーネントであり、Office アドインで使用できます。Fabric は Office デザイン言語を使用するため、Fabric の UX コンポーネントは Office に元々組み込まれているかのように自然に使うことができます。 

アドインを作成する場合は、Office UI Fabric を使用してユーザー エクスペリエンスを作成することをお勧めします。Office UI Fabric の使用は省略可能です。

次のセクションでは、Fabric を使用して要件を満たす方法について説明します。 

## <a name="use-fabric-core-icons-fonts-colors"></a>Fabric Core を使用する: アイコン、フォント、色
Fabric Core には、デザイン言語の基本的な要素 (アイコン、色、タイプ、グリッドなど) が含まれます。Fabric Core は独立したフレームワークです。Fabric React と Fabric JS は、どちらも Fabric Core を使用します。

Fabric Core の使用を開始するには

1. ページの HTML に CDN 参照を追加します。  

    ```html
    <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/css/fabric.min.css">
    ```   
    
2. Fabric のアイコンとフォントを使用します。 

    Fabric のアイコンを使用するには、ページに "i" 要素を含め、適切なクラスを参照します。アイコンのサイズは、フォント サイズを変更することで制御できます。たとえば、次のコードは、themePrimary (#0078d7) 色を使用する特大の表アイコンを作成する方法を示しています。 
   
    ```html
    <i class="ms-Icon ms-font-xl ms-Icon--Table ms-fontColor-themePrimary"></i>
    ```

    その他の Office UI Fabric で使用可能なアイコンを見つけるには、「[アイコン](https://dev.office.com/fabric#/styles/icons)」ページの検索機能を使用します。アドインで使用するアイコンを検索するときには、アイコン名の先頭に `ms-Icon--` を追加してください。 

    Office UI Fabric で使用可能なフォントのサイズと色については、「[文字体裁](https://dev.office.com/fabric#/styles/typography)」および「[色](https://dev.office.com/fabric#/styles/colors)」を参照してください。
 
## <a name="use-fabric-components"></a>Fabric コンポーネントを使用する 
Fabric には、次のタイプのコンポーネントを含む、さまざまな UX コンポーネントが用意されています。これらを使用してアドインを作成できます。

- 入力コンポーネント - 例: ボタン、チェック ボックス、および切り替え
- ナビゲーション コンポーネント - 例: Pivot Breadcrumb
- 通知コンポーネント - 例: MessageBar および Callout  

すべての Fabric コンポーネントがアドインでの使用に適しているわけではありません。このセクションで推奨されるコンポーネントの使用方法については、ガイダンスが用意されています。たとえば、アドインで Fabric のボタンを使用するためのガイダンスについては、「[ボタン](button.md)」を参照してください。 

アドインの作成には、Angular や React など別の JavaScript フレームワークも使用できます。フレームワークで Fabric コンポーネントを使用するには、次のリソースを参照してください。

|**フレームワーク**|**例**|
|:------------|:----------|
|**JavaScript のみ** (フレームワークでない)|[Office アドインで Office UI Fabric JS を使用する](using-office-ui-fabric-js.md)。|
|**React**|[Office アドインで Office UI Fabric React を使用する](using-office-ui-fabric-react.md )|
|**Angular**| Angular 1.5 ディレクティブのコミュニティ プロジェクトである「[ngOfficeUIFabric](http://ngofficeuifabric.com/)」と、「[Fabric コンポーネントと Angular 2 コンポーネントとのラッピングについて検討する](../develop/add-ins-with-angular2.md#consider-wrapping-fabric-components-with-angular-components)」を参照してください。|