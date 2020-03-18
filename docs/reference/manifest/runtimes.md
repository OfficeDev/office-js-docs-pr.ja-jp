---
title: マニフェストファイル内のランタイム (プレビュー)
description: Runtime 要素は、アドインのランタイムを指定します。
ms.date: 02/21/2020
localization_priority: Normal
ms.openlocfilehash: 5797aa78ae3667461de48de481ff44f14c307ced
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720422"
---
# <a name="runtimes-element-preview"></a>ランタイム要素 (プレビュー)

[!include[Running custom functions in browser runtime note](../../includes/excel-shared-runtime-preview-note.md)]

アドインのランタイムを指定し、カスタム関数、リボンボタン、および作業ウィンドウを使用して同じ JavaScript ランタイムを使用できるようにします。 マニフェストファイル内`<Host>`の要素の子。 詳細については、「[共有 JavaScript ランタイムを使用するように Excel アドインを構成する](../../excel/configure-your-add-in-to-use-a-shared-runtime.md)」を参照してください。

**アドインの種類:** 作業ウィンドウ

> [!IMPORTANT]
> 共有ランタイムは現在プレビュー段階であり、Windows 上の Excel でのみ使用できます。 プレビュー機能を試すには、 [Office Insider](https://insider.office.com/)に参加する必要があります。

## <a name="syntax"></a>構文

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a>含まれる場所 
[Host](./host.md)

## <a name="child-elements"></a>子要素

|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
|  **ランタイム**     | はい |  アドインのランタイム。

## <a name="see-also"></a>関連項目

- [ランタイム](runtime.md)
