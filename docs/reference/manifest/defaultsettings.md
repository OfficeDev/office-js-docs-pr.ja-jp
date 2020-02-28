---
title: マニフェスト ファイルの DefaultSettings 要素
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 824c575b39a99c6028ffd603390d2b41ee0ad7dd
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324885"
---
# <a name="defaultsettings-element"></a>DefaultSettings 要素

コンテンツまたは作業ウィンドウ アドインの既定のソースの場所と他の既定の設定を指定します。

**アドインの種類:** コンテンツ、作業ウィンドウ

## <a name="syntax"></a>構文

```XML
<DefaultSettings>
  ...
</DefaultSettings>
```

## <a name="contained-in"></a>含まれる場所

[OfficeApp](officeapp.md)

## <a name="can-contain"></a>含めることができるもの

|**Element**|**コンテンツ**|**メール**|**TaskPane**|
|:-----|:-----|:-----|:-----|
|[SourceLocation](sourcelocation.md)|x||x|
|[RequestedWidth](requestedwidth.md)|x|||
|[RequestedHeight](requestedheight.md)|x|||

## <a name="remarks"></a>注釈

**DefaultSettings**要素のソースの場所とその他の設定は、コンテンツアドインと作業ウィンドウアドインにのみ適用されます。メールアドインの場合は、 [formsettings](formsettings.md)要素に、ソースファイルとその他の既定の設定の既定の場所を指定します。

