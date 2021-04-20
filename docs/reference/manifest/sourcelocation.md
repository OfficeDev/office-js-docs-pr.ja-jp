---
title: マニフェスト ファイルの SourceLocation 要素
description: SourceLocation 要素は、Office アドインのソースファイルの場所を指定します。
ms.date: 05/12/2020
localization_priority: Normal
ms.openlocfilehash: 447adb7df7d0c59305fe5046357959fcd7824735
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641404"
---
# <a name="sourcelocation-element"></a>SourceLocation 要素

Office アドインのソースファイルの場所を、1 ~ 2018 文字の長さの URL として指定します。 ソースの場所はファイル パスではなく、HTTPS アドレスにする必要があります。

**アドインの種類:** コンテンツ、作業ウィンドウ、メール

## <a name="syntax"></a>構文

```XML
<SourceLocation DefaultValue="string" />
```

## <a name="contained-in"></a>含まれる場所

- [DefaultSettings](defaultsettings.md) (コンテンツ アドインおよび作業ウィンドウ アドイン)
- [FormSettings](formsettings.md) (メール アドイン)
- [Extensionpoint](extensionpoint.md) (コンテキストおよび launchevent (プレビュー) メールアドイン)

## <a name="can-contain"></a>含めることができるもの

[Override](override.md)

## <a name="attributes"></a>属性

|属性|型|必須|説明|
|:-----|:-----|:-----|:-----|
|DefaultValue|URL|必須|[DefaultLocale](defaultlocale.md) 要素に指定されるロケール用に、この設定の既定値を指定します。|
