---
title: マニフェスト ファイルの AppDomains 要素
description: Office アドインがページの読み込みに使用する`SourceLocation`要素に指定されているドメインに加えて、すべてのドメインを一覧表示します。
ms.date: 07/03/2019
localization_priority: Normal
ms.openlocfilehash: f60579d773e81a7e8006bafcf1c151874af42aeb
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720702"
---
# <a name="appdomains-element"></a>AppDomains 要素

Office アドインがページの読み込みに使用する`SourceLocation`要素に指定されているドメインに加えて、すべてのドメインを一覧表示します。 また、アドイン内の Iframe から Office .js API 呼び出しを行うことができる信頼されたドメインも一覧表示されます。 追加の各ドメインに、AppDomain 要素を指定します。

 **アドインの種類:** コンテンツ、作業ウィンドウ、メール

## <a name="syntax"></a>構文

```XML
<AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
</AppDomains>
```

> [!IMPORTANT]
> すべての **AppDomain** 要素の値には、プロトコル (例: `<AppDomain>https://myappdomain<AppDomain>`) が含まれている必要があります。

## <a name="contained-in"></a>含まれる場所

[OfficeApp](officeapp.md)

## <a name="can-contain"></a>含めることができるもの

[AppDomain](appdomain.md)

## <a name="remarks"></a>解説

アドインは、既定では [SourceLocation](sourcelocation.md) 要素で指定されたものと同じ場所のドメインのページを読み込みます。 アドインと同じドメインにないページを読み込む場合は、**AppDomains** 要素と **AppDomain** 要素を使用してドメインを指定します。 この要素は空にできません。
