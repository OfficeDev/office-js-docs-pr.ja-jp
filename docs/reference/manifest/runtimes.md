---
title: マニフェスト ファイル内のランタイム
description: Runtimes 要素は、アドインのランタイムを指定します。
ms.date: 05/14/2021
localization_priority: Normal
ms.openlocfilehash: e5ec70449d3984671d507131ac8d4fc0f7617cdcda1ad8f99b4f4bf52773aded
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2021
ms.locfileid: "57091630"
---
# <a name="runtimes-element"></a>Runtimes 要素

アドインのランタイムを指定します。 要素の [`<Host>`](host.md) 子。

> [!NOTE]
> Windows で Office で実行する場合、マニフェスト内に要素を持つアドインは、それ以外の場合と同じ Webview コントロールで必ずしも `<Runtimes>` 実行されるとは限りません。 Windows および Office のバージョンでどの webview コントロールが通常使用されるのかを決定する方法の詳細については、「Office アドインで使用されるブラウザー」を[参照](../../concepts/browsers-used-by-office-web-add-ins.md)してください。webView2 (Chromium ベース) で Microsoft Edge を使用する場合に説明されている条件が満たされている場合、アドインは要素を持っているかどうかに応じ、そのブラウザーを使用します `<Runtimes>` 。 ただし、これらの条件が満たされない場合、要素を持つアドインは、Windows またはバージョンに関係なく、常に Internet Explorer 11 Windows を `<Runtimes>` Microsoft 365します。

**アドインの種類:** 作業ウィンドウ, メール

[!include[Runtimes support](../../includes/runtimes-note.md)]

## <a name="syntax"></a>構文

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a>含まれる場所

[Host](host.md)

## <a name="child-elements"></a>子要素

|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
| [ランタイム](runtime.md) | 必要 |  アドインのランタイム。 **重要**: 現時点では、1 つの要素のみを定義 `<Runtime>` できます。 |

## <a name="see-also"></a>関連項目

- [ランタイム](runtime.md)
- [Office アドインを構成して共有 JavaScript ランタイムを使用する](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [イベント ベースのOutlook用にアドインを構成する](../../outlook/autolaunch.md)
