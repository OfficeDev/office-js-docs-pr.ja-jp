---
title: マニフェスト ファイル内のカスタム関数の SourceLocation 要素
description: Excel でカスタム関数によって使用される Script または Page 要素が必要とするリソースの場所を定義します。
ms.date: 02/02/2022
ms.localizationpriority: medium
---

# <a name="sourcelocation-element-custom-functions"></a>SourceLocation 要素 (カスタム関数)

カスタム関数で使用される **Script** 要素または **Page** 要素で必要なリソースの場所をExcel。

> [!IMPORTANT]
> この記事では、Page 要素または Script 要素の子である **SourceLocation** のみを **参照** します。 基本 [マニフェストの SourceLocation](sourcelocation.md) 要素の詳細については、「 **SourceLocation** 」を参照してください。

**アドインの種類:** カスタム関数

**次の VersionOverrides スキーマでのみ有効です**。

- Taskpane 1.0

詳細については、「Version [overrides in the manifest」を参照してください](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

**次の要件セットに関連付けられている**。

- [CustomFunctionsRuntime 1.1](../requirement-sets/custom-functions-requirement-sets.md)

## <a name="attributes"></a>属性

| 属性 | 必須 | 説明                                                                          |
|-----------|----------|--------------------------------------------------------------------------------------|
| resid     | はい      | マニフェストの **Resources** セクションで定義される URL リソースの名前。 32 文字以内で指定できます。 |

## <a name="child-elements"></a>子要素

なし

## <a name="example"></a>例

```xml
<SourceLocation resid="pageURL"/>
```
