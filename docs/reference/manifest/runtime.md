---
title: マニフェストファイル内のランタイム (プレビュー)
description: ''
ms.date: 02/21/2020
localization_priority: Normal
ms.openlocfilehash: dd51c5b317700f92ee74c94835e68523371789f8
ms.sourcegitcommit: 153576b1efd0234c6252433e22db213238573534
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/07/2020
ms.locfileid: "42561829"
---
# <a name="runtime-element-preview"></a>Runtime 要素 (プレビュー)

[!include[Running custom functions in browser runtime note](../../includes/excel-shared-runtime-preview-note.md)]

[`<Runtimes>`](runtimes.md)要素の子要素。 この要素は、リボン、作業ウィンドウ、およびカスタム関数がすべて同じランタイムで実行されるように、共有された JavaScript ランタイムを使用するようにアドインを構成します。 詳細については、「[共有 JavaScript ランタイムを使用するように Excel アドインを構成する](../../excel/configure-your-add-in-to-use-a-shared-runtime.md)」を参照してください。

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

- [ランタイム](runtimes.md)

## <a name="attributes"></a>属性

|  属性  |  必須  |  説明  |
|:-----|:-----|:-----|
|  **lifetime = "long"**  |  はい  | Excel アドインの`long`共有ランタイムを常に使用する場合は、必ず指定する必要があります。 |
|  **resid**  |  はい  | アドインの HTML ページの URL の場所を指定します。 は`resid` 、 `Resources`要素内`id`の`Url`要素の属性と一致している必要があります。 |

## <a name="see-also"></a>関連項目

- [ランタイム](runtimes.md)
