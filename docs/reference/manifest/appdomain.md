---
title: マニフェスト ファイルの AppDomain 要素
description: アドインで使用される追加のドメインを指定します。 Office によって信頼される必要があります。
ms.date: 06/12/2020
localization_priority: Normal
ms.openlocfilehash: ae49944afceada559b39353cd119e26a21fd3d15
ms.sourcegitcommit: 9eed5201a3ef556f77ba3b6790f007358188d57d
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/17/2020
ms.locfileid: "44778649"
---
# <a name="appdomain-element"></a>AppDomain 要素

[SourceLocation 要素](sourcelocation.md)で指定されているものに加えて、Office が信頼する必要がある追加のドメインを指定します。 ドメインの指定には、次のような影響があります。

- これにより、ドメイン内のページ、ルート、またはその他のリソースを、デスクトップの Office プラットフォーム上のアドインのルート作業ウィンドウで直接開くことができます。 (Web 上の Office、または IFrame でリソースを開くために**AppDomain**でドメインを指定する必要はありません)。または、[ダイアログ API](../../develop/dialog-api-in-office-add-ins.md)で開いたダイアログでリソースを開く必要はありません。
- これにより、ドメイン内のページは、アドイン内の Iframe から Office.js API 呼び出しを実行できるようになります。

**アドインの種類:** コンテンツ、作業ウィンドウ、メール

## <a name="syntax"></a>構文

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> 1. **AppDomain** 要素の値には、プロトコル (例: `<AppDomain>https://myappdomain.com</AppDomain>`) が含まれている必要があります。
> 2. ドメインの明示的なポートがある場合は、そのポートを含めます (例: `<AppDomain>https://myappdomain.com:9999</AppDomain>` )。
> 3. サブドメインを信頼する必要がある場合は、それを含めます (例: `<AppDomain>https://mysubdomain.myappdomain.com</AppDomain>` )。 サブドメイン `mysubdomain.mydomain.com` と `mydomain.com` ドメインが異なる。 両方を信頼する必要がある場合は、どちらも別個の**AppDomain**要素にする必要があります。
> 4. [SourceLocation 要素](sourcelocation.md)で指定されたドメインと同じドメインを一覧表示することはできません。誤解を招く可能性があります。 特に、を開発する場合は、 `localhost` 用の**AppDomain**要素を作成する必要はありません `localhost` 。
> 5. ドメインを超える URL のセグメントは含めないでください。 たとえば、ページの完全な URL を含めないでください。
> 6. 値には、末尾にスラッシュ "/" を付け*ない*でください。

## <a name="contained-in"></a>含まれる場所

[AppDomains](appdomains.md)

## <a name="remarks"></a>注釈

詳細については、「[Office アドインの XML マニフェスト](../../develop/add-in-manifests.md)」をご覧ください。
