---
title: Office アドインでの Office UI Fabric
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 7b1e4a9c377c9a60195a51115d7f275603f1ca5a
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2018
ms.locfileid: "23944035"
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

    その他の Office UI Fabric で使用可能なアイコンを見つけるには、「[アイコン](https://developer.microsoft.com/fabric#/styles/icons)」ページの検索機能を使用します。アドインで使用するアイコンを検索するときには、アイコン名の先頭に `ms-Icon--` を追加してください。 

    Office UI Fabric で使用可能なフォントのサイズと色については、「[文字体裁](https://developer.microsoft.com/fabric#/styles/typography)」および「[色](https://developer.microsoft.com/fabric#/styles/colors)」を参照してください。
 
## <a name="use-fabric-components"></a>Fabric コンポーネントを使用する 
Fabric には、次のタイプのコンポーネントを含む、さまざまな UX コンポーネントが用意されています。これらを使用してアドインを作成できます。

- 入力コンポーネント - 例: ボタン、チェック ボックス、および切り替え
- ナビゲーション コンポーネント - 例: ピボットおよび階層リンク
- 通知コンポーネントの MessageBar や吹き出しなど  

すべての Fabric コンポーネントがアドインでの使用を推奨しているわけではありません。アドインでの使用を推奨する Fabric React UX コンポーネントのリストを以下に示します。

- [階層リンク](https://developer.microsoft.com/fabric#/components/breadcrumb)
- [ボタン](https://developer.microsoft.com/fabric#/components/button)
- [チェックボックス](https://developer.microsoft.com/fabric#/components/checkbox)
- [選択肢グループ](https://developer.microsoft.com/fabric#/components/choicegroup)
- [ドロップダウン](https://developer.microsoft.com/fabric#/components/dropdown)
- [ラベル](https://developer.microsoft.com/fabric#/components/label)
- [リスト](https://developer.microsoft.com/fabric#/components/list)
- [コアドキュメント](https://developer.microsoft.com/fabric#/components/pivot)
- [TextField](https://developer.microsoft.com/fabric#/components/textfield)
- [切り替え](https://developer.microsoft.com/fabric#/components/toggle)

アドインの作成には、Angular や React など別の JavaScript フレームワークも使用できます。フレームワークで Fabric コンポーネントを使用するには、次のリソースを参照してください。

|**フレームワーク**|**例**|
|:------------|:----------|
|**応答**|[Office アドインで Office UI Fabric React を使用する](using-office-ui-fabric-react.md )|
|**角度**| Angular 1.5 ディレクティブのコミュニティ プロジェクトである「[ngOfficeUIFabric](http://ngofficeuifabric.com/)」と、「[Fabric コンポーネントと Angular 2 コンポーネントとのラッピングについて検討する](../develop/add-ins-with-angular2.md#consider-wrapping-fabric-components-with-angular-components)」を参照してください。|
