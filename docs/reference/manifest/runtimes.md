---
title: マニフェストファイル内のランタイム
description: ''
ms.date: 01/06/2020
localization_priority: Normal
ms.openlocfilehash: ec2b85a92325eb4e36c61f731369ec54d44ef169
ms.sourcegitcommit: 0dacbe7c80ed387099e3ec21e151f8990b181ede
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/13/2020
ms.locfileid: "41111178"
---
# <a name="runtimes-element"></a>ランタイム要素

この機能はプレビュー段階です。 アドインのランタイムを指定し、カスタム関数と作業ウィンドウでグローバルデータを共有して、関数呼び出しを相互に行うことができるようにします。 マニフェストファイルの`<Host>`要素に従う必要があります。

**アドインの種類:** 作業ウィンドウ

## <a name="syntax"></a>構文

```XML
<Runtimes>
    <Runtime resid="TaskPaneAndCustomFunction.Url" lifetime="long" />
</Runtimes>
```

## <a name="child-elements"></a>子要素

|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
|  **ランタイム**     | はい |  アドインのランタイム。多くの場合、Excel カスタム関数で使用されます。

## <a name="see-also"></a>関連項目

-[ランタイム](runtimes.md)
