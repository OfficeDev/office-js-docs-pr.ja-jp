---
title: マニフェスト ファイルの DefaultSettings 要素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 0c109d5d893cf9d3502f1cbf1724007f01e623e6
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433755"
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

|**要素**|**Content**|**Mail**|**TaskPane**|
|:-----|:-----|:-----|:-----|
|[SourceLocation](sourcelocation.md)|x||x|
|[RequestedWidth](requestedwidth.md)|x|||
|[RequestedHeight](requestedheight.md)|x|||

## <a name="remarks"></a>注釈

**DefaultSettings** 要素のソースの場所と他の設定が適用されるのは、コンテンツ アドインと作業ウィンドウ アドインのみです。メール アドインの場合は、ソース ファイルの既定の場所とその他の既定の設定を [FormSettings](formsettings.md) 要素に指定します。

