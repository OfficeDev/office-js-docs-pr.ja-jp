---
title: マニフェスト ファイルの AppDomains 要素
description: Office アドインが使用する要素で指定されたドメインに加えて、すべてのドメインを一覧表示し `SourceLocation` 、Office。
ms.date: 06/12/2020
localization_priority: Normal
ms.openlocfilehash: 55401d62e88cc1f2d67d13de0997a40db7a3f6b0c2f8997aa1b976962c8c797f
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2021
ms.locfileid: "57096534"
---
# <a name="appdomains-element"></a>AppDomains 要素

要素で指定されたドメインに加えて、Office アドインが使用し、Office によって信頼される必要があるドメイン `SourceLocation` を一覧表示します。 これにより、ドメイン内のページでアドイン内の IFrame Office.js API を呼び出し、その他の効果があります。 追加の各ドメインに、**AppDomain** 要素を指定します。

 **アドインの種類:** コンテンツ、作業ウィンドウ、メール

## <a name="syntax"></a>構文

```XML
<AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
</AppDomains>
```

> [!IMPORTANT]
> **AppDomain** 要素の値に制限があります。 詳細については [、「AppDomain」を参照してください](appdomain.md)。

## <a name="contained-in"></a>含まれる場所

[OfficeApp](officeapp.md)

## <a name="can-contain"></a>含めることができるもの

[AppDomain](appdomain.md)

## <a name="remarks"></a>解説

アドインは、既定では [SourceLocation](sourcelocation.md) 要素で指定されたものと同じ場所のドメインのページを読み込みます。 この要素は空にできません。
