---
title: マニフェストファイル内のランタイム
description: ''
ms.date: 01/24/2020
localization_priority: Normal
ms.openlocfilehash: 6682887935ee6894b5a311ad519408067452bb23
ms.sourcegitcommit: 72d719165cc2b64ac9d3c51fb8be277dfde7d2eb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/25/2020
ms.locfileid: "41554007"
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

- [ランタイム](runtime.md)
