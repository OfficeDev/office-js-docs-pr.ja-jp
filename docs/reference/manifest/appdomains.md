---
title: マニフェスト ファイルの AppDomains 要素
description: '`SourceLocation`Office アドインが使用する、office によって信頼される必要がある、要素で指定されているドメインに加えて、すべてのドメインを一覧表示します。'
ms.date: 06/12/2020
localization_priority: Normal
ms.openlocfilehash: 751e4ad2ffa5fd50739a855fad48964473b154f1
ms.sourcegitcommit: 9eed5201a3ef556f77ba3b6790f007358188d57d
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/17/2020
ms.locfileid: "44778656"
---
# <a name="appdomains-element"></a>AppDomains 要素

Office `SourceLocation` アドインが使用し、office によって信頼されるようにする必要がある、要素で指定されているドメインに加えて、すべてのドメインを一覧表示します。 これにより、ドメイン内のページは、アドイン内の Iframe から Office.js Api を呼び出すことができるようになり、他の効果があります。 追加の各ドメインに、**AppDomain** 要素を指定します。

 **アドインの種類:** コンテンツ、作業ウィンドウ、メール

## <a name="syntax"></a>構文

```XML
<AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
</AppDomains>
```

> [!IMPORTANT]
> **AppDomain**要素の値には、いくつかの制限があります。 詳細については、「 [AppDomain](appdomain.md)」を参照してください。

## <a name="contained-in"></a>含まれる場所

[OfficeApp](officeapp.md)

## <a name="can-contain"></a>含めることができるもの

[AppDomain](appdomain.md)

## <a name="remarks"></a>解説

アドインは、既定では [SourceLocation](sourcelocation.md) 要素で指定されたものと同じ場所のドメインのページを読み込みます。 この要素は空にできません。
