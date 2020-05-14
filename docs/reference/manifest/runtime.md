---
title: マニフェストファイル内のランタイム
description: Runtime 要素は、アドインが、リボン、作業ウィンドウ、およびカスタム関数に対して共有 JavaScript ランタイムを使用するように構成します。
ms.date: 05/11/2020
localization_priority: Normal
ms.openlocfilehash: c5c7356f9985ca7b5972068629b0587f8916348e
ms.sourcegitcommit: 682d18c9149b1153f9c38d28e2a90384e6a261dc
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/13/2020
ms.locfileid: "44217761"
---
# <a name="runtime-element"></a>Runtime 要素

要素の子要素 [`<Runtimes>`](runtimes.md) 。 この要素は、リボン、作業ウィンドウ、およびカスタム関数がすべて同じランタイムで実行されるように、共有された JavaScript ランタイムを使用するようにアドインを構成します。 詳細については、「[共有 JavaScript ランタイムを使用するように Excel アドインを構成する](../../excel/configure-your-add-in-to-use-a-shared-runtime.md)」を参照してください。

**アドインの種類:** 作業ウィンドウ

## <a name="syntax"></a>構文

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a>含まれる場所

- [ランタイム](runtimes.md)

## <a name="attributes"></a>属性

|  属性  |  必須  |  説明  |
|:-----|:-----|:-----|
|  **lifetime = "long"**  |  はい  | Excel アドインの共有ランタイムを常に使用する場合は、必ず指定する必要があり `long` ます。 |
|  **resid**  |  はい  | アドインの HTML ページの URL の場所を指定します。 は、 `resid` `id` 要素内の要素の属性と一致している必要があり `Url` `Resources` ます。 |

## <a name="see-also"></a>関連項目

- [ランタイム](runtimes.md)
