---
title: マニフェスト ファイルの Script 要素
description: Script 要素は、カスタム関数がカスタム 関数で使用するスクリプト設定をExcel。
ms.date: 02/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: 0f32314912dd66d8578750bf4818af8483c8ef36
ms.sourcegitcommit: 61c183a5d8a9d889b6934046c7e4a217dc761b80
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/16/2022
ms.locfileid: "62855542"
---
# <a name="script-element"></a>Script 要素

Excel でカスタム関数によって使用されるスクリプトの設定を定義します。

**アドインの種類:** カスタム関数

**次の VersionOverrides スキーマでのみ有効です**。

- Taskpane 1.0

詳細については、「Version [overrides in the manifest」を参照してください](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

**次の要件セットに関連付けられている**。

- [CustomFunctionsRuntime 1.1](../requirement-sets/custom-functions-requirement-sets.md)

## <a name="attributes"></a>属性

なし

## <a name="child-elements"></a>子要素

|要素  |  必須  |  説明  |
|:-----|:-----|:-----|
|  [SourceLocation](customfunctionssourcelocation.md)  |  はい  | カスタム関数によって使用される JavaScript ファイルのリソース ID を持つ文字列。|

## <a name="example"></a>例

```xml
<Script>
    <SourceLocation resid="scriptURL" />
    <!-- The Script element is not used in the Developer Preview. -->
</Script>
```
