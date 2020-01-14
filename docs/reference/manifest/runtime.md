---
title: マニフェストファイル内のランタイム
description: ''
ms.date: 01/06/2020
localization_priority: Normal
ms.openlocfilehash: 68def44ba74733934198ac3b32fa1fe649156766
ms.sourcegitcommit: 0dacbe7c80ed387099e3ec21e151f8990b181ede
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/13/2020
ms.locfileid: "41111171"
---
# <a name="runtime-element"></a>Runtime 要素

この機能はプレビュー段階です。 [`<Runtimes>`](runtime.md)要素の子要素。 この要素を使用すると、Excel カスタム関数とアドインの作業ウィンドウの間でのグローバルデータと関数呼び出しの共有が容易になります。 

## <a name="contained-in"></a>含まれる場所

-[ランタイム](runtimes.md)

**アドインの種類:** 作業ウィンドウ

## <a name="syntax"></a>構文

```XML
<Runtimes>
    <Runtime resid="TaskPaneAndCustomFunction.Url" lifetime="long" />
</Runtimes>
```

## <a name="attributes"></a>属性

|  属性  |  必須  |  説明  |
|:-----|:-----|:-----|
|  **lifetime = "long"**  |  はい  | アドインの作業ウィンドウが閉じているときに Excel カスタム関数が動作するようにする場合は、常に long として表示される必要があります。 |
|  **resid**  |  はい  | Excel カスタム関数で使用する場合、 `resid`はを`TaskPaneAndCustomFunction.Url`参照する必要があります。 |

## <a name="see-also"></a>関連項目

-[ランタイム](runtime.md)
