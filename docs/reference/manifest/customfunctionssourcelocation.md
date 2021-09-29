---
title: マニフェスト ファイル内のカスタム関数の SourceLocation 要素
description: Excel でカスタム関数によって使用される Script または Page 要素が必要とするリソースの場所を定義します。
ms.date: 09/24/2021
ms.localizationpriority: medium
ms.openlocfilehash: 5f2d881f31f4e46e7f5bb8ab30d78abd0e9b7200
ms.sourcegitcommit: 517786511749c9910ca53e16eb13d0cee6dbfee6
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/29/2021
ms.locfileid: "59990685"
---
# <a name="sourcelocation-element-custom-functions"></a>SourceLocation 要素 (カスタム関数)

Excel でカスタム関数によって使用される Script または Page 要素が必要とするリソースの場所を定義します。

**アドインの種類:** カスタム関数

## <a name="attributes"></a>属性

| 属性 | 必須 | 説明                                                                          |
|-----------|----------|--------------------------------------------------------------------------------------|
| resid     | はい      | マニフェストの &lt;Resources&gt; セクションで定義される URL リソースの名前。 32 文字以内で指定できます。 |

## <a name="child-elements"></a>子要素

なし

## <a name="example"></a>例

```xml
<SourceLocation resid="pageURL"/>
```
