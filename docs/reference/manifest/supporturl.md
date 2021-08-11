---
title: マニフェスト ファイルの SupportUrl 要素
description: SupportUrl 要素は、アドインのサポート情報を提供するページの URL を指定します。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 1d76afeaaceafc9e8786070338d69cea1b73635d20cd5a729d7e3d859b952494
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2021
ms.locfileid: "57096361"
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
