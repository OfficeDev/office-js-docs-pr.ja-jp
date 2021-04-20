---
title: Office アドインでの Office UI Fabric
description: アドインで UI Fabric コンポーネントをOfficeする方法の概要Office説明します。
ms.date: 2/09/2021
localization_priority: Normal
ms.openlocfilehash: 9799d98d795486203e4bcc23bffc043c2ead6e28
ms.sourcegitcommit: ccc0a86d099ab4f5ef3d482e4ae447c3f9b818a3
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/14/2021
ms.locfileid: "50237680"
---
# <a name="office-ui-fabric-in-office-add-ins"></a>Office アドインでの Office UI Fabric

Office UI Fabric は、ユーザー エクスペリエンスを構築するための JavaScript フロントエンド フレームワークOffice。 Fabric は、拡張や改訂が可能な視覚効果に焦点を合わせたコンポーネントであり、Office アドインで使用できます。 Fabric は Office デザイン言語を使用するため、Fabric の UX コンポーネントは Office に元々組み込まれているかのように自然に使うことができます。

アドインを作成する場合は、Office UI Fabric を使用してユーザー エクスペリエンスを作成することをお勧めします。Office UI Fabric の使用は省略可能です。

次のセクションでは、Fabric を使用して要件を満たす方法について説明します。

## <a name="use-fabric-core-icons-fonts-colors"></a>Fabric Core を使用する: アイコン、フォント、色

Fabric Core には、デザイン言語の基本的な要素 (アイコン、色、タイプ、グリッドなど) が含まれます。 Fabric Core は独立したフレームワークです。 Fabric Core は、Fabric React によって使用され、Fabric React に含まれます。

Fabric Core の使用を開始するには:

1. ページの HTML に CDN 参照を追加します。  

    ```html
    <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/9.6.1/css/fabric.min.css">
    ```

2. Fabric のアイコンとフォントを使用します。

    Fabric のアイコンを使用するには、ページに "i" 要素を含め、適切なクラスを参照します。アイコンのサイズは、フォント サイズを変更することで制御できます。たとえば、次のコードは、themePrimary (#0078d7) 色を使用する特大の表アイコンを作成する方法を示しています。

    ```html
    <i class="ms-Icon ms-font-xl ms-Icon--Table ms-fontColor-themePrimary"></i>
    ```

    その他の Office UI Fabric で使用可能なアイコンを見つけるには、「[アイコン](https://developer.microsoft.com/fabric#/styles/icons)」ページの検索機能を使用します。アドインで使用するアイコンを検索するときには、アイコン名の先頭に `ms-Icon--` を追加してください。

    Office UI Fabric で使用可能なフォントのサイズと色については、「[文字体裁](https://developer.microsoft.com/fabric#/styles/typography)」および「[色](https://developer.microsoft.com/fabric#/styles/colors)」を参照してください。

## <a name="use-fabric-components"></a>Fabric コンポーネントを使用する

Fabric は、アドインの構築に使用できるさまざまな UX コンポーネントを提供します。 すべてのファブリック コンポーネントが 1 つのアドインで使用されるとは予想されません。 シナリオとユーザー エクスペリエンスに最適なコンポーネントを決定します (たとえば、作業ウィンドウに [階層](https://developer.microsoft.com/fabric#/components/breadcrumb) リンクを適切に表示するのは難しい場合があります)。

アドインでの使用をお勧めする [一](https://developer.microsoft.com/fluentui#/controls/web) 般的な Fabric React UX コンポーネントの一覧を次に示します。

- [Button](https://developer.microsoft.com/fabric#/components/button)
- [Checkbox](https://developer.microsoft.com/fabric#/components/checkbox)
- [ChoiceGroup](https://developer.microsoft.com/fabric#/components/choicegroup)
- [Dropdown](https://developer.microsoft.com/fabric#/components/dropdown)
- [Label](https://developer.microsoft.com/fabric#/components/label)
- [List](https://developer.microsoft.com/fabric#/components/list)
- [Pivot](https://developer.microsoft.com/fabric#/components/pivot)
- [TextField](https://developer.microsoft.com/fabric#/components/textfield)
- [Toggle](https://developer.microsoft.com/fabric#/components/toggle)

アドインの作成には、Angular や React など別の JavaScript フレームワークも使用できます。フレームワークで Fabric コンポーネントを使用するには、次のリソースを参照してください。

|**フレームワーク**|**例**|
|:------------|:----------|
|**React**|[Office アドインで Office UI Fabric React を使用する](using-office-ui-fabric-react.md )|
|**Angular**| [Fabric コンポーネントと Angular 2 コンポーネントの折り返しを検討する](../develop/add-ins-with-angular2.md#consider-wrapping-fabric-components-with-angular-components)|
