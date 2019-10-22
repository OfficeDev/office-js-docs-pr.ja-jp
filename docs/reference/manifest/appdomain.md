---
title: マニフェスト ファイルの AppDomain 要素
description: ''
ms.date: 07/03/2019
localization_priority: Normal
ms.openlocfilehash: 2f65302d1ac3d85f2867cd13501bc67606cd00b5
ms.sourcegitcommit: b3996b1444e520b44cf752e76eef50908386ca26
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/21/2019
ms.locfileid: "35575640"
---
# <a name="appdomain-element"></a>AppDomain 要素

アドインウィンドウにページを読み込む追加のドメインを指定します。 また、アドイン内の Iframe から Office .js API 呼び出しを行うことができる信頼されたドメインも一覧表示されます。

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
