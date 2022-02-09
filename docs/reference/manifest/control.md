---
title: マニフェスト ファイルの Control 要素
description: アクションを実行するか、作業ウィンドウを起動するコントロールを定義します。
ms.date: 02/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: aa7ff9b0162070b378352ce187de15a34323b998
ms.sourcegitcommit: d01aa8101630031515bf27f14361c5a3062c3ec4
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/09/2022
ms.locfileid: "62467837"
---
# <a name="control-element"></a>Control 要素

アクションを実行するか、作業ウィンドウを起動するコントロールを定義します。 **Control** 要素は、[ボタン] または [メニュー] オプションのどちらかになります。 少なくとも 1 つの **Control** に 1 つの [Group](group.md) 要素を含む必要があります。

**アドインの種類:** 作業ウィンドウ, メール

**次の VersionOverrides スキーマでのみ有効です**。

- 作業ウィンドウ 1.0
- メール 1.0
- メール 1.1

詳細については、「Version [overrides in the manifest」を参照してください](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

**次の要件セットに関連付けられている**。

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md) (作業ウィンドウ アドインの場合)。
- 一部の子要素は、追加の要件セットに関連付けられる場合があります。

## <a name="attributes"></a>属性

|  属性  |  必須  |  説明  |
|:-----|:-----|:-----|
|**xsi:type**|はい|定義されているコントロールの型。 、、または`Button``Menu`を指定できます`MobileButton`。 |
|**id**|はい|コントロール要素の ID です。 最大で 125 文字です。 マニフェスト内のすべての **Control** 要素で一意である必要があります。|

> [!NOTE]
> **xsi:type** の `MobileButton` 値は、VersionOverrides スキーマ 1.1 で定義されます。 これは、[MobileFormFactor](mobileformfactor.md) 要素内に含まれる **Control** 要素にのみ当てはまります。

## <a name="child-elements"></a>子要素

有効な子要素は、 **xsi:type 属性の値によって異** なります。

- [Control 要素のボタンの種類](control-button.md)
- [Control 要素のメニューの種類](control-menu.md)
- [Control 要素の MobileButton 型](control-mobilebutton.md)
