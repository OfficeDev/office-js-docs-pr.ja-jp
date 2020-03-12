---
title: マニフェスト ファイルの AppDomain 要素
description: ''
ms.date: 07/03/2019
localization_priority: Normal
ms.openlocfilehash: da28b3b4dec5d669462a781db3c0628bd32c7182
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596789"
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

**AppDomain** 要素は、[SourceLocation](sourcelocation.md) 要素で指定したドメイン以外のものを追加指定するために使用する必要があります。 詳細については、「[Office アドイン XML マニフェスト](../../develop/add-in-manifests.md)」を参照してください。
