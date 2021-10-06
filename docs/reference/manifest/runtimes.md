---
title: マニフェスト ファイル内のランタイム
description: Runtimes 要素は、アドインのランタイムを指定します。
ms.date: 09/28/2021
ms.localizationpriority: medium
ms.openlocfilehash: 758bb7b830009d6691190a0279440a52da724624
ms.sourcegitcommit: 489befc41e543a4fb3c504fd9b3f61322134c1ef
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/06/2021
ms.locfileid: "60138605"
---
# <a name="runtimes-element"></a>Runtimes 要素

アドインのランタイムを指定します。 要素の [`<Host>`](host.md) 子。

> [!NOTE]
> Windows で Office で実行する場合、マニフェスト内に要素を持つアドインは、それ以外の場合と同じ Webview コントロールで必ずしも `<Runtimes>` 実行されるとは限りません。 Windows および Office のバージョンでどの webview コントロールが通常使用されるのかを決定する方法の詳細については、「Office アドインで使用されるブラウザー」を[参照](../../concepts/browsers-used-by-office-web-add-ins.md)してください。webView2 (Chromium ベース) で Microsoft Edge を使用する場合に説明されている条件が満たされている場合、アドインは要素を持っているかどうかに応じ、そのブラウザーを使用します `<Runtimes>` 。 ただし、これらの条件が満たされない場合、要素を持つアドインは、Windows またはバージョンに関係なく、常に Internet Explorer 11 Windows を `<Runtimes>` Microsoft 365します。

**アドインの種類:** 作業ウィンドウ, メール

**次の VersionOverrides スキーマでのみ有効です**。

 - 作業ウィンドウ 1.0
 - メール 1.1

詳細については、「マニフェストの [バージョンオーバーライド」を参照してください](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

**次の要件セットに関連付けられている**。

- [SharedRuntime 1.1](../requirement-sets/shared-runtime-requirement-sets.md) (作業ウィンドウ アドインで使用する場合のみ)。

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
| [Runtime](runtime.md) | はい |  アドインのランタイム。 **重要**: 現時点では、1 つの要素のみを定義 `<Runtime>` できます。 |

## <a name="see-also"></a>関連項目

- [Runtime](runtime.md)
- [Office アドインを構成して共有 JavaScript ランタイムを使用する](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [イベント ベースのOutlook用にアドインを構成する](../../outlook/autolaunch.md)
