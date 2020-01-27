---
title: マニフェストファイル内のランタイム
description: ''
ms.date: 01/24/2020
localization_priority: Normal
ms.openlocfilehash: 8fbad8276b3e1d64a6c443cf57d498597d729282
ms.sourcegitcommit: 72d719165cc2b64ac9d3c51fb8be277dfde7d2eb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/25/2020
ms.locfileid: "41554000"
---
# <a name="runtime-element"></a>Runtime 要素

この機能はプレビュー段階です。 [`<Runtimes>`](runtimes.md)要素の子要素。 この要素を使用すると、Excel カスタム関数とアドインの作業ウィンドウの間でのグローバルデータと関数呼び出しの共有が容易になります。

**アドインの種類:** 作業ウィンドウ

## <a name="syntax"></a>構文

```XML
<Runtimes>
    <Runtime resid="TaskPaneAndCustomFunction.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a>含まれる場所

- [ランタイム](runtimes.md)

## <a name="attributes"></a>属性

|  属性  |  必須  |  説明  |
|:-----|:-----|:-----|
|  **lifetime = "long"**  |  はい  | アドインの作業ウィンドウが閉じているときに Excel カスタム関数が動作するようにする場合は、常に long として表示される必要があります。 |
|  **resid**  |  はい  | Excel カスタム関数で使用する場合、 `resid`はを`TaskPaneAndCustomFunction.Url`参照する必要があります。 |

## <a name="see-also"></a>関連項目

- [ランタイム](runtimes.md)
