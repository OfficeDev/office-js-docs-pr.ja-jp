---
title: マニフェスト ファイルの SourceLocation 要素
description: SourceLocation 要素は、アドインのソース ファイルOffice指定します。
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: 6830a26cf192802c97c486511695b4ace35ac8263cfcd30ceaf71398f0d83a07
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2021
ms.locfileid: "57095009"
---
# <a name="sourcelocation-element"></a>SourceLocation 要素

1 ~ 2018 文字の URL として、Officeアドインのソース ファイルの場所を指定します。 ソースの場所はファイル パスではなく、HTTPS アドレスにする必要があります。

**アドインの種類:** コンテンツ、作業ウィンドウ、メール

## <a name="syntax"></a>構文

```XML
<SourceLocation DefaultValue="string" />
```

## <a name="contained-in"></a>含まれる場所

- [DefaultSettings](defaultsettings.md) (コンテンツ アドインおよび作業ウィンドウ アドイン)
- [FormSettings](formsettings.md) (メール アドイン)
- [ExtensionPoint](extensionpoint.md) (コンテキスト メール アドインと LaunchEvent メール アドイン)

## <a name="can-contain"></a>含めることができるもの

[Override](override.md)

## <a name="attributes"></a>属性

|属性|型|必須|説明|
|:-----|:-----|:-----|:-----|
|DefaultValue|URL|必須|[DefaultLocale](defaultlocale.md) 要素に指定されるロケール用に、この設定の既定値を指定します。|
