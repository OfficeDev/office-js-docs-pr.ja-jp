---
title: マニフェスト ファイルの AppDomain 要素
description: アドインで使用され、ユーザーが信頼する必要がある追加のドメインをOffice。
ms.date: 06/12/2020
localization_priority: Normal
ms.openlocfilehash: ae49944afceada559b39353cd119e26a21fd3d15
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/08/2021
ms.locfileid: "58938117"
---
# <a name="appdomain-element"></a>AppDomain 要素

SourceLocation 要素で指定されたドメインOffice、信頼する必要がある追加のドメイン[を指定します](sourcelocation.md)。 ドメインを指定すると、次の効果があります。

- これにより、ドメイン内のページ、ルート、または他のリソースを、デスクトップ プラットフォーム上のアドインのルート作業ウィンドウで直接開Officeできます。 **(AppDomain** でドメインを指定しても、Office on the web または IFrame でリソースを開く必要はありません。また [、Dialog API](../../develop/dialog-api-in-office-add-ins.md)で開いたダイアログでリソースを開く必要もありません。
- これにより、ドメイン内のページがアドインOffice.js IFrames から API 呼び出しを実行できます。

**アドインの種類:** コンテンツ、作業ウィンドウ、メール

## <a name="syntax"></a>構文

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> 1. **AppDomain** 要素の値には、プロトコル (例: `<AppDomain>https://myappdomain.com</AppDomain>`) が含まれている必要があります。
> 2. ドメインに明示的なポートがある場合は、それを含める (たとえば `<AppDomain>https://myappdomain.com:9999</AppDomain>` )。
> 3. サブドメインを信頼する必要がある場合は、サブドメインを含める (例: `<AppDomain>https://mysubdomain.myappdomain.com</AppDomain>` )。 サブドメインであり `mysubdomain.mydomain.com` 、 `mydomain.com` 異なるドメインです。 両方を信頼する必要がある場合は、両方とも別々の **AppDomain 要素に含む必要** があります。
> 4. [SourceLocation](sourcelocation.md)要素で指定されたドメインと同じドメインを一覧に表示すると、効果が得らなく、誤解を招く可能性があります。 特に、上で開発する場合 `localhost` は **、AppDomain** 要素を作成する必要があります `localhost` 。
> 5. ドメインを越える URL のセグメントを含めない。 たとえば、ページの完全な URL を含めない。
> 6. 値 *に* "/"というスラッシュを付け込む必要があります。

## <a name="contained-in"></a>含まれる場所

[AppDomains](appdomains.md)

## <a name="remarks"></a>注釈

詳細については、「[Office アドインの XML マニフェスト](../../develop/add-in-manifests.md)」をご覧ください。
