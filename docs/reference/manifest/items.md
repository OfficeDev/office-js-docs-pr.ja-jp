---
title: マニフェスト ファイル内の Items 要素
description: メニューの項目を指定します。
ms.date: 02/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: 2249bc55db662a36cf3986ebb0b90353237d4985
ms.sourcegitcommit: d01aa8101630031515bf27f14361c5a3062c3ec4
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/09/2022
ms.locfileid: "62467927"
---
# <a name="items-element"></a>Items 要素

メニューの項目を指定します。

**アドインの種類:** 作業ウィンドウ, メール

**次の VersionOverrides スキーマでのみ有効です**。

- 作業ウィンドウ 1.0
- メール 1.0
- メール 1.1

詳細については、「Version [overrides in the manifest」を参照してください](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

**次の要件セットに関連付けられている**。

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md) 親 **VersionOverrides が** Taskpane 1.0 と入力されている場合。
- 親 **VersionOverrides が Mail** 1.0 と入力されている場合のメールボックス [1.3](../../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md)。
- 親 **VersionOverrides が Mail** 1.1 と入力されている場合のメールボックス [1.5](../../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)。

## <a name="syntax"></a>構文

```XML
<Items>
...  
</Items>  
```

## <a name="contained-in"></a>含まれる場所

[Menu 型のコントロール要素](control-menu.md)

## <a name="must-contain"></a>含める必要があるもの

[Item](item.md)

## <a name="examples"></a>例

例については、「Control [of type Menu」を参照してください](control-menu.md)。