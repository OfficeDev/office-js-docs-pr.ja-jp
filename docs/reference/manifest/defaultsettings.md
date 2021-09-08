---
title: マニフェスト ファイルの DefaultSettings 要素
description: コンテンツまたは作業ウィンドウ アドインの既定のソースの場所と他の既定の設定を指定します。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: a9711fb44390bcbda8979b8018eed1318c5579bc
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/08/2021
ms.locfileid: "58936349"
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

|要素|コンテンツ|メール|TaskPane|
|:-----|:-----|:-----|:-----|
|[SourceLocation](sourcelocation.md)|x||x|
|[RequestedWidth](requestedwidth.md)|x|||
|[RequestedHeight](requestedheight.md)|x|||

## <a name="remarks"></a>注釈

**DefaultSettings** 要素のソースの場所と他の設定は、コンテンツ アドインと作業ウィンドウ アドインにのみ適用されます。メール アドインの場合は [、FormSettings](formsettings.md)要素でソース ファイルの既定の場所と他の既定の設定を指定します。
