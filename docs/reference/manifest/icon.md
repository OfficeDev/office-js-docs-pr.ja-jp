---
title: マニフェスト ファイルの Icon 要素
description: ボタン または メニュー コントロールの Image 要素を定義します。
ms.date: 10/30/2021
ms.localizationpriority: medium
ms.openlocfilehash: 54ae88e5dceeffa244780764711b263ceabd828d
ms.sourcegitcommit: 23ce57b2702aca19054e31fcb2d2f015b4183ba1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/02/2021
ms.locfileid: "60681727"
---
# <a name="icon-element"></a>Icon 要素

[ボタン](control.md#button-control) または [メニュー](control.md#menu-dropdown-button-controls) コントロールの **Image** 要素を定義します。

## <a name="attributes"></a>属性

|  属性  |  必須  |  説明  |
|:-----|:-----|:-----|
|  **xsi:type**  |  いいえ  | 定義されているアイコンの型。これは、モバイル フォーム ファクターのアイコンにのみ適用されます [MobileFormFactor](mobileformfactor.md) 要素に含まれる **Icon** 要素は、この属性を `bt:MobileIconList` に設定する必要があります。 |

## <a name="child-elements"></a>子要素

|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
|  [Image](#image)        | はい |   使用するイメージの resid         |

### <a name="image"></a>Image

ボタンの画像です。 **resid 属性** は 32 文字以内で **、Resources** 要素の **Images** 要素の Image 要素の **id** 属性の値に [設定する必要](resources.md)があります。 **size** 属性は、画像のサイズをピクセル単位で示します。 他に 5 つのサイズ (20、24、40、48、64 ピクセル) がサポートされていますが、3 つの画像のサイズ (16、32、80 ピクセル) を必ず指定します。

```xml
<Icon>
  <bt:Image size="16" resid="blue-icon-16" />
  <bt:Image size="32" resid="blue-icon-32" />
  <bt:Image size="80" resid="blue-icon-80" />
</Icon>
```

> [!IMPORTANT]
> この画像がアドインの代表的なアイコンである場合は、「Create effective listings in [AppSource](/office/dev/store/create-effective-office-store-listings#create-an-icon-for-your-add-in) and within Office サイズと他の要件」を参照してください。

## <a name="additional-requirements-for-mobile-form-factors"></a>モバイル フォーム ファクターの追加要件

親 **Icon** 要素が、[MobileFormFactor](mobileformfactor.md) 要素の子孫である場合は、必要な最小サイズが若干異なります。 マニフェストで、最小サイズを 25、32、および 48 ピクセルに指定する必要があります。 指定するサイズは、`scale`、`1` または `2` に設定された `3` 属性で必ずそれぞれ 3 回ずつ表示されます。 この属性は `UIScreen.scale` 、iOS デバイスのプロパティを指定します。 詳細については、「scale」を [参照してください](https://developer.apple.com/documentation/uikit/uiscreen/1617836-scale)。

```xml
<Icon xsi:type="bt:MobileIconList">
  <bt:Image resid="blue-icon-16-1" size="25" scale="1" />
  <bt:Image resid="blue-icon-16-2" size="25" scale="2" />
  <bt:Image resid="blue-icon-16-3" size="25" scale="3" />
  <bt:Image resid="blue-icon-32-1" size="32" scale="1" />
  <bt:Image resid="blue-icon-32-2" size="32" scale="2" />
  <bt:Image resid="blue-icon-32-3" size="32" scale="3" />
  <bt:Image resid="blue-icon-80-1" size="48" scale="1" />
  <bt:Image resid="blue-icon-80-2" size="48" scale="2" />
  <bt:Image resid="blue-icon-80-3" size="48" scale="3" />
</Icon>
```
