---
title: マニフェストファイル内のランタイム
description: ''
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: 945a30527632b23a594d7bfb82cec94e74754249
ms.sourcegitcommit: dc42e0276007f8ab006028b9cd0cc1526c1bd100
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/14/2020
ms.locfileid: "41120636"
---
# <a name="runtime-element"></a>Runtime 要素

この機能はプレビュー段階です。 [`<Runtimes>`](runtime.md)要素の子要素。 この要素を使用すると、Excel カスタム関数とアドインの作業ウィンドウの間でのグローバルデータと関数呼び出しの共有が容易になります。

**アドインの種類:** 作業ウィンドウ

## <a name="syntax"></a>構文

```XML
<Runtimes>
    <Runtime resid="TaskPaneAndCustomFunction.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a>含まれる場所

-[ランタイム](runtimes.md)

## <a name="attributes"></a>属性

|  属性  |  必須  |  説明  |
|:-----|:-----|:-----|
|  **lifetime = "long"**  |  はい  | アドインの作業ウィンドウが閉じているときに Excel カスタム関数が動作するようにする場合は、常に long として表示される必要があります。 |
|  **resid**  |  はい  | Excel カスタム関数で使用する場合、 `resid`はを`TaskPaneAndCustomFunction.Url`参照する必要があります。 |

## <a name="see-also"></a>関連項目

-[ランタイム](runtime.md)
