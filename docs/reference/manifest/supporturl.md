---
title: マニフェスト ファイルの SupportUrl 要素
description: SupportUrl 要素は、アドインのサポート情報を提供するページの URL を指定します。
ms.date: 10/09/2018
ms.localizationpriority: medium
ms.openlocfilehash: 2ea515aa61ed5bf9e22d6316a76fa4b5e51493f3
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2021
ms.locfileid: "59151242"
---
# <a name="supporturl-element"></a>SupportUrl 要素

アドインのサポート情報を提供するページの URL を指定します。

## <a name="syntax"></a>構文

```XML
<OfficeApp>
...
  <IconUrl DefaultValue="https://contoso.com/assets/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="https://contoso.com/assets/hi-res-icon.png"/>
  
  
  <SupportUrl DefaultValue="https://contoso.com/support " />
  
  
  <AppDomains>
  ...
  </AppDomains>
...
</OfficeApp>
```

## <a name="contained-in"></a>含まれる場所

[OfficeApp](officeapp.md)

## <a name="can-contain"></a>含めることができるもの

|  要素 | 必須 | 説明  |
|:-----|:-----|:-----|
|  [Override](override.md)   | なし | 追加のロケール URL の設定を指定します。 |

## <a name="attributes"></a>属性

|属性|型|必須|説明|
|:-----|:-----|:-----|:-----|
|DefaultValue|URL|必須|この設定の既定値を指定します。この値は、[DefaultLocale](defaultlocale.md) 要素に指定されるロケールを対象としています。|
