---
title: マニフェスト ファイルの AppDomain 要素
description: ''
ms.date: 03/21/2019
localization_priority: Normal
ms.openlocfilehash: 8216603c87a7dcafde84d25a82f068c9aa86ed96
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870409"
---
# <a name="appdomain-element"></a>AppDomain 要素

アドイン ウィンドウにページを読み込むために使用される追加のドメインを指定します。

**アドインの種類:** コンテンツ、作業ウィンドウ、メール

## <a name="syntax"></a>構文

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> 1. **AppDomain** 要素の値には、プロトコル (例: `<AppDomain>https://myappdomain</AppDomain>`) が含まれている必要があります。
> 2. 値には、末尾にスラッシュ "/" を付け*ない*でください。

## <a name="contained-in"></a>含まれる場所

[AppDomains](appdomains.md)

## <a name="remarks"></a>解説

**AppDomain** 要素は、[SourceLocation](sourcelocation.md) 要素で指定したドメイン以外のものを追加指定するために使用する必要があります。 詳細については、「[Office アドイン XML マニフェスト](/office/dev/add-ins/develop/add-in-manifests)」を参照してください。
