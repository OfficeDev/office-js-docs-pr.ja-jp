---
title: マニフェスト ファイルの SourceLocation 要素
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 7544e2bae480b9431c8912533ea1b761132a355e
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32451977"
---
# <a name="sourcelocation-element"></a>SourceLocation 要素

Office アドインのソース ファイルの場所を、1 から 2018 文字までの長さの URL として指定します。ソースの場所はファイル パスではなく、HTTPS アドレスにする必要があります。

**アドインの種類:** コンテンツ、作業ウィンドウ、メール

## <a name="syntax"></a>構文

```XML
<SourceLocation DefaultValue="string" />
```

## <a name="contained-in"></a>含まれる場所

- [DefaultSettings](defaultsettings.md) (コンテンツ アドインおよび作業ウィンドウ アドイン)
- [FormSettings](formsettings.md) (メール アドイン)
- [ExtensionPoint](extensionpoint.md) (コンテキスト メール アドイン)

## <a name="can-contain"></a>含めることができるもの

[Override](override.md)

## <a name="attributes"></a>属性

|**属性**|**型**|**必須**|**説明**|
|:-----|:-----|:-----|:-----|
|DefaultValue|URL|必須|[DefaultLocale](defaultlocale.md) 要素に指定されるロケール用に、この設定の既定値を指定します。|
